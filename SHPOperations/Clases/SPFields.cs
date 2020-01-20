using Microsoft.SharePoint;

namespace SHPOperations.Clases
{
    public enum TypeColumn
    {
        LineOfText = 0,
        Multiline = 1,
        Choice = 2,
        Number = 3,
        Lookup = 4,
        UsersAndGroups = 5
    }

    public enum TypeUserSelectionMode
    {
        PeopleOnly = 0,
        PeopleAndGroups = 1
    }

    /// <summary>
    /// Configuración general para la columna
    /// </summary>
    public partial class SPFields
    {
        public TypeColumn Column { get; set; }
        public string ListName { get; set; }
        public string ColumnName { get; set; }
        public string Description { get; set; }
        public string Type { get; set; }
        public SPFieldText FieldText { get; set; }
        public SPFieldMultiText FieldMultiText { get; set; }
        public SPFieldChoice FieldChoice { get; set; }
        public SPFieldNumber FieldNumber { get; set; }
        public SPFieldLookup FieldLookup { get; set; }
        public SPFieldUsersAndGroup FieldUserGroup { get; set; }
    }

    /// <summary>
    /// Configuración para las columnas de tipo UNA LINEA DE TEXTO
    /// </summary>
    public partial class SPFieldText
    {
        public bool ContainInformation { get; set; }
        public bool UniqueValues { get; set; }
        public int MaxLength { get; set; }
        public string DefaultValue { get; set; }
    }

    /// <summary>
    /// Configuración para las columnas de tipo VARIAS LINEAS DE TEXTO
    /// </summary>
    public partial class SPFieldMultiText
    {
        public bool ContainInformation { get; set; }
        public int NumLines { get; set; }
        public bool RichText { get; set; }
    }

    /// <summary>
    /// Configuración para las columnas de tipo ELECCION (menú para elegir)
    /// </summary>
    public partial class SPFieldChoice
    {
        public bool ContainInformation { get; set; }
        public bool UniqueValues { get; set; }
        public string[] Options { get; set; }
        public string DefaultValue { get; set; }
    }

    /// <summary>
    /// Configuración para las columnas de tipo NUMERICO
    /// </summary>
    public partial class SPFieldNumber
    {
        public double MinValue { get; set; }
        public double MaxValue { get; set; }
        public bool ShowAsPercentaje { get; set; }
    }

    /// <summary>
    /// Configuración para las columnas de tipo LOOKUP
    /// </summary>
    public partial class SPFieldLookup
    {
        public bool ContainInformation { get; set; }
        public string LookupList { get; set; }
        public string LookupField { get; set; }
    }

    /// <summary>
    /// Configuración para las columnas de tipo FIELDUSER
    /// </summary>
    public partial class SPFieldUsersAndGroup
    {
        public bool ContainInformation { get; set; }
        public TypeUserSelectionMode UserSelectionMode { get; set; }
        public string ShowField { get; set; }
        public bool MultiUser { get; set; }
        public string UsersFromGroup { get; set; }
    }
}
