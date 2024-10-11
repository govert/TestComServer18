using System.Runtime.InteropServices;
using ExcelDna.Integration;

namespace TestComServer18
{
    [ComVisible(false)]
    public class ExcelAddin : IExcelAddIn
    {
        public void AutoOpen()
        {
            ExcelDna.ComInterop.ComServer.DllRegisterServer();
        }
        public void AutoClose()
        {
            ExcelDna.ComInterop.ComServer.DllUnregisterServer();
        }
    }
}