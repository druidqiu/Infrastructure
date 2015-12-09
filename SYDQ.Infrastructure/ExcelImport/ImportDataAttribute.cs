using System;

namespace SYDQ.Infrastructure.ExcelImport
{
    public class ImportDataAttribute : Attribute
    {
        public ImportDataAttribute(string description)
        {
            Description = description;
        }

        public string Description { get; set; }
        public bool Required { get; set; }
        public string EmptyErrorMessageFormat { get; set; }
        public string DataTypeErrorMessageFormat { get; set; }
    }
}
