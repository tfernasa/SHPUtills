using System.Collections.Generic;

namespace SHPOperations.Clases
{
    public enum TypeColumn
    {
        LineOfText = 0,
        Multiline = 1,
        Choice = 2
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
}
