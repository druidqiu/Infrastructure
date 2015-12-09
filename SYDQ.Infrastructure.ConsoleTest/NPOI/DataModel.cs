using SYDQ.Infrastructure.ExcelExport;
using SYDQ.Infrastructure.ExcelImport;

namespace SYDQ.Infrastructure.ConsoleTest.NPOI
{
    [Export("用户信息")]
    public class ExportDataModel
    {
        [Export("用户名")]
        public string UserName { get; set; }

        [Export("用户密码")]
        public string Password { get; set; }

        [Export("序号")]
        public int Index { get; set; }
    }

    [ImportData("用户信息")]
    public class ImportDataModel
    {
        [ImportData("用户名称", Required = true,
            RequiredErrorMessageFormat="{0}不能为空。", 
            DataTypeErrorMessageFormat = "")]
        public string UserName { get; set; }
        [ImportData("密码", Required = true, DataTypeErrorMessageFormat = "")]
        public string Password { get; set; }
        [ImportData("序号", Required = true, DataTypeErrorMessageFormat = "{0}只能是数字。")]
        public int Index { get; set; }
    }
}
