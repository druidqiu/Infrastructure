using System;

namespace SYDQ.Infrastructure.NPOI
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
