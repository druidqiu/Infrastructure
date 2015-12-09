using System;

namespace SYDQ.Infrastructure.ExcelExport
{
    public class ExportAttribute : Attribute
    {
        public ExportAttribute(string description)
        {
            Description = description;
        }

        public string Description { get; set; }
    }
}
