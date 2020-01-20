using System;

namespace SHPOperations.Clases
{
    public enum TypeError
    {
        Warning = 0,
        Error = 1,
    }

    public partial class ClassError
    {
        public TypeError Error { get; set; }
        public string Source { get; set; }
        public string ErrorMessage { get; set; }
        public string AditionalInfo { get; set; }
        public DateTime DateError { get; set; }
    }
}
