using SYDQ.Infrastructure.ExcelExport;
using SYDQ.Infrastructure.ExcelImport;

namespace SYDQ.Infrastructure.ConsoleTest.NPOI
{
    [ExportDescription("用户信息")]
    public class ExportDataModel
    {
        [ExportDescription("用户名")]
        public string UserName { get; set; }

        [ExportDescription("用户密码")]
        public string Password { get; set; }

        [ExportDescription("序号")]
        public int Index { get; set; }
    }

    [ImportData("用户信息")]
    public class ImportDataModel
    {
        [ImportData("用户名称", Required = true,
            EmptyErrorMessageFormat="{0}不能为空。", 
            DataTypeErrorMessageFormat = "")]
        public string UserName { get; set; }
        [ImportData("密码", Required = true, DataTypeErrorMessageFormat = "")]
        public string Password { get; set; }
        [ImportData("序号", Required = true, DataTypeErrorMessageFormat = "{0}只能是数字。")]
        public int Index { get; set; }
    }
}
