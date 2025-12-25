using ExcelDna.Integration;
using ExcelDna.IntelliSense;

namespace Lisha.ExcelAddins.ExcelDna
{
    public class IntelliSenseAddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            IntelliSenseServer.Install();
        }
        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();
        }
    }
}
