using System;

namespace SYDQ.Infrastructure.ExcelExport
{
    public class ExportDescriptionAttribute : Attribute
    {
        public ExportDescriptionAttribute(string description)
        {
            Description = description;
        }

        public string Description { get; set; }
    }
}
