using SYDQ.Infrastructure.NPOI;

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
}
