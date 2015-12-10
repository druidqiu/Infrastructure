using System;

namespace SYDQ.Infrastructure.ExcelImport
{
    public class ImportAttribute : Attribute
    {
        public ImportAttribute(string description)
        {
            Description = description;
        }

        public string Description { get; set; }
        public bool Required { get; set; }
        public string RequiredErrorMessageFormat { get; set; }
        public string DataTypeErrorMessageFormat { get; set; }
    }
}
