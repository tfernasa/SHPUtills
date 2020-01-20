using System;
using System.Net;
using System.Security;
using System.Collections.Generic;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Client;
using SHPOperations.Clases;
using SP = Microsoft.SharePoint;

namespace SHPOperations
{
    /// <summary>
    /// Clase para gestionar las operaciones con SHP (Creación de librerías, etc...)
    /// </summary>
    public class SHPExecOperations : IDisposable
    {
        //--------------------------------------------------------------------
        #region variables y constantes (privadas)
        bool disposed = false;
        ClientContext site = null;
        const string schemaXML = "<Field Type='{0}' DisplayName='{1}' Name='{1}' />";
        const string schemaXMLLookUp = "<Field Type='{0}' DisplayName='{1}' Required='{2}' List='{3}' ShowField = '{4}' StaticName = '{4}' Name = '{4}' />";
        #endregion
        //--------------------------------------------------------------------

        //--------------------------------------------------------------------
        #region Propiedades
        public string SiteSHP { get; set; }
        public string SiteDOMAIN { get; set; }
        public string SiteUSER { get; set; }
        public string SitePWD { get; set; }
        public string ListName { get; set; }
        public bool IsSHPOnline { get; set; }
        public bool ConectionSiteOK { get; set; }
        public List<ClassError> SHPExecOperationsErrors { get; set; }
        public List<SPFields> SPFields { get; set; }
        #endregion
        //--------------------------------------------------------------------

        //--------------------------------------------------------------------
        #region Constructores y destructores de la clase
        /// <summary>
        /// Constructor de la clase
        /// </summary>
        public SHPExecOperations()
        {
            SHPExecOperationsErrors = new List<ClassError>();
            SPFields = new List<SPFields>();
            site = null;
            ListName = "";
            this.SiteSHP = SPContext.Current.Web.Url;
            ConectionSiteOK = ConnectToSite();
        }

        /// <summary>
        /// Constructor de la clase
        /// </summary>
        /// <param name="siteSHP">Url del site Sharepoint</param>
        /// <param name="siteUSER">Opcional -> Credenciales (usuario) Ej: TFERNASA o USERSAD\TFERNASA</param>
        /// <param name="sitePWD">Opcional -> Credenciales (Pwd)</param>
        /// <param name="isSHPOnline">Opcional -> Indica si las credenciales son para SHP Online (office 365) o no (versiones anteriores shp)</param>
        public SHPExecOperations(string siteSHP, string siteUSER = "", string sitePWD = "", bool isSHPOnline = false)
        {
            //Establecemos - Inicializamos valores...
            SHPExecOperationsErrors = new List<ClassError>();
            SPFields = new List<SPFields>();
            site = null;
            this.IsSHPOnline = isSHPOnline;
            this.SiteSHP = siteSHP.Trim();
            this.SiteUSER = siteUSER.Trim();
            this.SitePWD = sitePWD.Trim();
            ListName = "";
            this.SiteDOMAIN = "";
            ConectionSiteOK = ConnectToSite();
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                //Objetos a liberar / eliminar
                site = null;
                SHPExecOperationsErrors = null;
            }

            disposed = true;
        }
        #endregion
        //--------------------------------------------------------------------

        //--------------------------------------------------------------------
        #region Procedimientos y funciones varios (PUBLIC)
        /// <summary>
        /// Crea una nueva lista en el site. 
        /// </summary>
        /// <param name="Title">Título / Nombre de la lista</param>
        /// <param name="TemplateType">Tipo de lista (Genérica (100) por defecto)</param>
        /// <param name="Description">Descripcion para la lista a crear</param>
        /// <param name="SetListName">Indica si se ha de establecer el nombre de la lista a una propiedad de la clase para poder realizar, a posteriori, mas operaciones como añadir campos</param>
        /// <returns></returns>
        public bool CreateList(string title, int templateType = 100, string description = "", bool setListName = false)
        {
            bool result;
            SPFields.Clear();

            if (ConectionSiteOK)
            {
                ListCreationInformation CreationInfo = new ListCreationInformation
                {
                    Title = title.Trim(),
                    Description = description.Trim(),
                    TemplateType = templateType
                };

                try
                {
                    List newList = site.Web.Lists.Add(CreationInfo);
                    site.Load(newList);
                    site.ExecuteQuery();

                    result = true;
                }
                catch (Exception ex)
                {
                    SHPExecOperationsErrors.Add(GetInfoError(ex, string.Format("Alta de lista {0}", title), TypeError.Error));
                    result = false;
                }
            }
            else
            {
                SHPExecOperationsErrors.Add(GetInfoError("Alta de nueva lista (CreateList)", "No se ha establecido conexión con el site", string.Format("Alta de lista {0}", title), TypeError.Warning));
                result = false;
            }

            if ((result) && (setListName))
            {
                ListName = title.Trim();
            }

            return result;
        }

        //--------------------------------------------------------------------
        #region Crear nuevos campos
        /// <summary>
        /// Añade una nueva columna de tipo "Linea de Texto"
        /// </summary>
        /// <param name="columnName">Nombre de la columna</param>
        /// <param name="containInformation">Idica si la columna debe de contener información (campo requerido)</param>
        /// <param name="uniqueValues">Indica si en la coolumna se deben aplicar valores únicos </param>
        /// <param name="maxLength">Indica el nº máximo de caracteres para la columna (Entre 1 y 255)</param>
        /// <param name="defaultValue">Opcional - Indicia el valor por defecto de la columna (solo texto, valor calculado)</param>
        /// <param name="listName">Opcional - Nombre de la lista en la que crear la nueva columna. Si no se indica valor, se tomará el nombre definido en la propiedad ListName</param>
        /// <param name="description">Opcional - Descripción del campo</param>
        public void AddNewColumn(string columnName, bool containInformation, bool uniqueValues , int maxLength, string defaultValue = "", string listName = "", string description = "")
        {
            listName = SetListName(listName);

            if ((maxLength <=0) || (maxLength > 255))
            {
                maxLength = 255;
            }

            SPFields field = new SPFields
            {
                Column = TypeColumn.LineOfText,
                ListName = listName.Trim(),
                ColumnName = columnName.Trim(),
                Description = description.Trim(),
                Type = "Text",
                FieldText = new Clases.SPFieldText
                {
                    ContainInformation = containInformation,
                    UniqueValues = uniqueValues,
                    MaxLength = maxLength,
                    DefaultValue = defaultValue.Trim()
                }
            };

            SPFields.Add(field);
        }

        /// <summary>
        /// Añade una nueva columna de tipoo "Varias líneas de texto"
        /// </summary>
        /// <param name="columnName">Nombre de la columna</param>
        /// <param name="containInformation">Idica si la columna debe de contener información (campo requerido)</param>
        /// <param name="numLines">Indica el nº de líneas de texto para la columna</param>
        /// <param name="richText">Indica el tipo de texto que se permite (FALSE = Texto sin formato / TRUE = Texto enriquecido)</param>
        /// <param name="listName">Opcional - Nombre de la lista en la que crear la nueva columna. Si no se indica valor, se tomará el nombre definido en la propiedad ListName</param>
        /// <param name="description">Opcional - Descripción del campo</param>
        public void AddNewColumn(string columnName, bool containInformation, int numLines, bool richText, string listName = "", string description = "")
        {
            listName = SetListName(listName);

            if (numLines<=0)
            {
                numLines = 6;
            }

            SPFields field = new SPFields
            {
                Column = TypeColumn.Multiline,
                ListName = listName.Trim(),
                ColumnName = columnName.Trim(),
                Description = description.Trim(),
                Type = "Note",
                FieldMultiText = new SPFieldMultiText
                {
                    ContainInformation = containInformation,
                    NumLines = numLines,
                    RichText = richText
                }
            };

            SPFields.Add(field);
        }

        /// <summary>
        /// Añade una nueva columna de tipo "Elección (menú para elegir)"
        /// </summary>
        /// <param name="columnName">Nombre de la columna</param>
        /// <param name="containInformation">Idica si la columna debe de contener información (campo requerido)</param>
        /// <param name="uniqueValues">Indica si en la coolumna se deben aplicar valores únicos </param>
        /// <param name="options">Lista de opciones para rellenar la columna</param>
        /// <param name="typeChoice">Tipo de opciones (Menu desplegable, botones o casillas)</param>
        /// <param name="defaultValue">Opcional - Valor predeterminado (debe de coincidir con uno de los elementos que componen el parametro 'options')</param>
        /// <param name="listName">Opcional - Nombre de la lista en la que crear la nueva columna. Si no se indica valor, se tomará el nombre definido en la propiedad ListName</param>
        /// <param name="description">Opcional - Descripción del campo</param>
        public void AddNewColumn(string columnName, bool containInformation, bool uniqueValues, string[] options,  string defaultValue = "", string listName = "", string description = "")
        {
            listName = SetListName(listName);

            SPFields field = new SPFields
            {
                Column = TypeColumn.Choice,
                ListName = listName.Trim(),
                ColumnName = columnName.Trim(),
                Description = description.Trim(),
                Type = "Choice",
                FieldChoice = new Clases.SPFieldChoice
                {
                    ContainInformation = containInformation,
                    UniqueValues = uniqueValues,
                    Options = options,
                    DefaultValue = defaultValue.Trim()
                }
            };

            SPFields.Add(field);
        }

        /// <summary>
        /// Añade una nueva columna de tipo "Numero"
        /// </summary>
        /// <param name="columnName">Nombre de la columna</param>
        /// <param name="showAsPercentaje">Indica si el campo numerico es un campo para representar porcentajes (TRUE) o no (FALSE)</param>
        /// <param name="minValue">Valor numérico mínimo para el campo (null si no desea indicar valor mínimo)</param>
        /// <param name="MaxValue">Valor numérico máximo para el campo (null si no desea indicar valor máximo)</param>
        /// <param name="listName">Opcional - Nombre de la lista en la que crear la nueva columna. Si no se indica valor, se tomará el nombre definido en la propiedad ListName</param>
        /// <param name="description">Opcional - Descripción del campo</param>
        public void AddNewColumn(string columnName, bool showAsPercentaje, Nullable<double> minValue, Nullable<double> maxValue, string listName = "", string description = "")
        {
            listName = SetListName(listName);

            SPFields field = new SPFields
            {
                Column = TypeColumn.Number,
                ListName = listName.Trim(),
                ColumnName = columnName.Trim(),
                Description = description.Trim(),
                Type = "Number",
                FieldNumber = new Clases.SPFieldNumber
                {
                    ShowAsPercentaje = showAsPercentaje
                }
            };

            //Campos nulos
            if (double.TryParse(minValue.ToString(), out double Value1))
            {
                field.FieldNumber.MinValue = Value1;
            }

            if (double.TryParse(maxValue.ToString(), out double Value2))
            {
                field.FieldNumber.MaxValue = Value2;
            }

            if ((field.FieldNumber.MinValue != 0) && (field.FieldNumber.MaxValue != 0) && (field.FieldNumber.MinValue > field.FieldNumber.MaxValue))
            {
                //Si se da la casualidad de que el valor minimo indicado es mayor al valor máximo, anulamos ambos valores (quitamos minimo y maximo)
                field.FieldNumber.MinValue = 0;
                field.FieldNumber.MaxValue = 0;
            }

            SPFields.Add(field);
        }

        /// <summary>
        /// Añade una nueva columna de tipo "Lookup"
        /// </summary>
        /// <param name="columnName">Nombre de la columna</param>
        /// <param name="containInformation">Idica si la columna debe de contener información (campo requerido)</param>
        /// <param name="listName">Opcional - Nombre de la lista en la que crear la nueva columna. Si no se indica valor, se tomará el nombre definido en la propiedad ListName</param>
        /// <param name="lookupList">Nombre de la lista con la cual se desea vincular el campo de busqueda</param>
        /// <param name="lookupField">Nombre del campo a mostrar</param>
        ///<param name="description">Opcional - Descripción del campo</param>
        public void AddNewColumn(string columnName, bool containInformation, string lookupList, string lookupField, string listName = "", string description = "")
        {
            listName = SetListName(listName);

            SPFields field = new SPFields
            {
                Column = TypeColumn.Lookup,
                ListName = listName.Trim(),
                ColumnName = columnName.Trim(),
                Description = description.Trim(),
                Type = "Lookup",
                FieldLookup = new Clases.SPFieldLookup
                {
                    ContainInformation = containInformation,
                    LookupList = lookupList,
                    LookupField = lookupField
                }
            };

            SPFields.Add(field);
        }

        /// <summary>
        /// Lanza el proceso de creación de nuevas columnas
        /// </summary>
        /// <returns></returns>
        public bool CreateColumns()
        {
            bool result = true;
            List list = null;
            bool continueProcess = true;

            if (ConectionSiteOK)
            {
                if (SPFields.Count > 0)
                {
                    foreach (SPFields infoField in SPFields)
                    {
                        continueProcess = true;
                        try
                        {
                            list = site.Web.Lists.GetByTitle(infoField.ListName);
                        }
                        catch (Exception ex)
                        {
                            SHPExecOperationsErrors.Add(GetInfoError(ex, string.Format("Alta de nuevo campo {0}", infoField.ColumnName), TypeError.Error));
                            continueProcess = false;
                        }

                        if (continueProcess)
                        {
                            //Configuramos...
                            switch (infoField.Column)
                            {
                                case TypeColumn.LineOfText:
                                    //Linea de texto
                                    Field newField1 = list.Fields.AddFieldAsXml(string.Format(schemaXML, infoField.Type, infoField.ColumnName), true, AddFieldOptions.AddToDefaultContentType);
                                    newField1.StaticName = infoField.ColumnName;
                                    newField1.Description = infoField.Description;

                                    if (infoField.FieldText != null)
                                    {
                                        FieldText fldText = site.CastTo<FieldText>(newField1);
                                        fldText.Required = infoField.FieldText.ContainInformation;
                                        fldText.EnforceUniqueValues = infoField.FieldText.UniqueValues;
                                        fldText.MaxLength = infoField.FieldText.MaxLength;
                                        fldText.DefaultValue = infoField.FieldText.DefaultValue;
                                        if (infoField.FieldText.UniqueValues)
                                        {
                                            fldText.Indexed = true;
                                        }
                                        fldText.Update();
                                    }
                                    else
                                    {
                                        continueProcess = false;
                                    }
                                    break;
                                case TypeColumn.Multiline:
                                    //Varias líneas de texto
                                    Field newField2 = list.Fields.AddFieldAsXml(string.Format(schemaXML, infoField.Type, infoField.ColumnName), true, AddFieldOptions.AddToDefaultContentType);
                                    newField2.StaticName = infoField.ColumnName;
                                    newField2.Description = infoField.Description;

                                    if (infoField.FieldMultiText!= null)
                                    {
                                        FieldMultiLineText fldMultiLine = site.CastTo<FieldMultiLineText>(newField2);
                                        fldMultiLine.Required = infoField.FieldMultiText.ContainInformation;
                                        fldMultiLine.NumberOfLines = infoField.FieldMultiText.NumLines;
                                        fldMultiLine.RichText = infoField.FieldMultiText.RichText;
                                        fldMultiLine.Update();
                                    }
                                    else
                                    {
                                        continueProcess = false;
                                    }
                                    break;
                                case TypeColumn.Choice:
                                    //Eleccion
                                    Field newField3 = list.Fields.AddFieldAsXml(string.Format(schemaXML, infoField.Type, infoField.ColumnName), true, AddFieldOptions.AddToDefaultContentType);
                                    newField3.StaticName = infoField.ColumnName;
                                    newField3.Description = infoField.Description;

                                    if (infoField.FieldChoice!=null)
                                    {
                                        FieldMultiChoice fldChoice = site.CastTo<FieldMultiChoice>(newField3);
                                        fldChoice.Required = infoField.FieldChoice.ContainInformation;
                                        fldChoice.EnforceUniqueValues = infoField.FieldChoice.UniqueValues;
                                        fldChoice.DefaultValue = infoField.FieldChoice.DefaultValue;
                                        if (infoField.FieldChoice.UniqueValues)
                                        {
                                            fldChoice.Indexed = true;
                                        }
                                        fldChoice.Choices = infoField.FieldChoice.Options;
                                        fldChoice.Update();
                                    }
                                    else
                                    {
                                        continueProcess = false;
                                    }
                                    break;
                                case TypeColumn.Number:
                                    //Numero
                                    Field newField4 = list.Fields.AddFieldAsXml(string.Format(schemaXML, infoField.Type, infoField.ColumnName), true, AddFieldOptions.AddToDefaultContentType);
                                    newField4.StaticName = infoField.ColumnName;
                                    newField4.Description = infoField.Description;

                                    if (infoField.FieldNumber != null)
                                    {
                                        FieldNumber fldNumber = site.CastTo<FieldNumber>(newField4);
                                        fldNumber.ShowAsPercentage = infoField.FieldNumber.ShowAsPercentaje;
                                        if (infoField.FieldNumber.MinValue != 0)
                                        {
                                            fldNumber.MinimumValue = infoField.FieldNumber.MinValue;
                                        }
                                        if (infoField.FieldNumber.MaxValue != 0)
                                        {
                                            fldNumber.MaximumValue = infoField.FieldNumber.MaxValue;
                                        }

                                        fldNumber.Update();
                                    }
                                    else
                                    {
                                        continueProcess = false;
                                    }
                                    break;
                                case TypeColumn.Lookup:
                                    if (infoField.FieldLookup != null)
                                    {
                                        string required = "FALSE";
                                        if (infoField.FieldLookup.ContainInformation)
                                        {
                                            required = "TRUE";
                                        }

                                        List sourceList = site.Web.Lists.GetByTitle(infoField.FieldLookup.LookupList);
                                        site.Load(sourceList);
                                        
                                        List destinationList = site.Web.Lists.GetByTitle(infoField.ListName);
                                        site.Load(destinationList);

                                        site.ExecuteQuery();

                                        Field LookUpField = destinationList.Fields.AddFieldAsXml(string.Format(schemaXMLLookUp, infoField.Type, infoField.ColumnName, required, sourceList.Id, infoField.FieldLookup.LookupField), true, AddFieldOptions.DefaultValue);
                                        LookUpField.Description = infoField.Description;
                                        LookUpField.Update();
                                    }
                                    else
                                    {
                                        continueProcess = false;
                                    }
                                    break;
                            }

                            //Añadimos...
                            if (continueProcess)
                            {
                                try
                                {
                                    site.ExecuteQuery();
                                }
                                catch (Exception ex)
                                {
                                    SHPExecOperationsErrors.Add(GetInfoError(ex, string.Format("Alta de nuevo campo {0}", infoField.ColumnName), TypeError.Error));
                                }
                            }
                        }

                        list = null;
                    }
                }
                else
                {
                    SHPExecOperationsErrors.Add(GetInfoError("Creación de columnas (CreateColumns)", "No se han indicado campos a crear", "", TypeError.Warning));
                    result = false;
                }
            }
            else
            {
                SHPExecOperationsErrors.Add(GetInfoError("Creación de columnas (CreateColumns)", "No se ha establecido conexión con el site", "", TypeError.Warning));
                result = false;
            }

            SPFields.Clear();

            return result;
        }
        #endregion
        //--------------------------------------------------------------------
        #endregion
        //--------------------------------------------------------------------

        //--------------------------------------------------------------------
        #region Procedimientos y funciones varios (PRIVATE)
        /// <summary>
        /// Conecta con un site SHP (crea el objeto)
        /// </summary>
        /// <returns></returns>
        private bool ConnectToSite()
        {
            bool testConnection = false;
            site = null;
            try
            {
                site = new ClientContext(SiteSHP);
                if (!string.IsNullOrWhiteSpace(SiteUSER) && !string.IsNullOrWhiteSpace(SitePWD))
                {
                    testConnection = true;
                    if (SiteUSER.Contains(@"\"))
                    {
                        SiteDOMAIN = SiteUSER.Substring(0, SiteUSER.IndexOf(@"\"));
                        SiteUSER = SiteUSER.Substring(SiteUSER.IndexOf(@"\") + 1);
                        site.Credentials = new NetworkCredential(SiteUSER, SitePWD, SiteDOMAIN);
                    }
                    else
                    {
                        if (IsSHPOnline)
                        {
                            site.Credentials = new SharePointOnlineCredentials(SiteUSER, GetPasswordSecurity(SitePWD));
                        }
                        else
                        {
                            site.Credentials = new NetworkCredential(SiteUSER, SitePWD);
                        }
                    }
                }
                else
                {
                    site.Credentials = new NetworkCredential();
                }

                if (testConnection)
                {
                    Web webTest = site.Web;
                    try
                    {
                        site.Load(webTest);
                        site.ExecuteQuery();
                        return true;
                    }
                    catch (Exception ex)
                    {
                        SHPExecOperationsErrors.Add(GetInfoError(ex, "Conexion a site - credenciales", Clases.TypeError.Error));
                        return false;
                    }
                }
                else
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                SHPExecOperationsErrors.Add(GetInfoError(ex, "Conexion a site", Clases.TypeError.Error));
                return false;
            }
        }

        private string SetListName(string listName)
        {
            if (string.IsNullOrEmpty(listName))
            {
                listName = this.ListName.Trim();
            }

            return listName.Trim();
        }

        /// <summary>
        /// Devuelve un objeto con toda la info de un error sucedido para agregar a la lista de errores (traza de errores)
        /// </summary>
        /// <param name="infoEx"></param>
        /// <param name="action"></param>
        /// <returns></returns>
        private ClassError GetInfoError(Exception infoEx, string action, TypeError error)
        {
            ClassError InfoError = new ClassError
            {
                Error = error,
                Source = infoEx.Source,
                ErrorMessage = infoEx.Message,
                AditionalInfo = string.Format("Acción que ha provocado el error: {0}", action),
                DateError = DateTime.Now
            };

            return InfoError;
        }

        /// <summary>
        /// Devuelve un objeto con toda la info de un error sucedido para agregar a la lista de errores (traza de errores)
        /// </summary>
        /// <param name="Source"></param>
        /// <param name="Message"></param>
        /// <param name="Action"></param>
        /// <param name="Error"></param>
        /// <returns></returns>
        private ClassError GetInfoError(string source, string message, string action, TypeError error)
        {
            ClassError InfoError = new ClassError
            {
                Error = error,
                Source = source,
                ErrorMessage = message,
                AditionalInfo = string.Format("Acción que ha provocado el error: {0}", action),
                DateError = DateTime.Now
            };

            return InfoError;
        }

        private SecureString GetPasswordSecurity(string sitePWD)
        {
            SecureString securePassword = new SecureString();

            char[] charArr = sitePWD.ToCharArray();

            foreach (char ch in charArr)
            {
                securePassword.AppendChar(ch);
            }

            return securePassword;
        }
        #endregion
        //--------------------------------------------------------------------
    }
}

