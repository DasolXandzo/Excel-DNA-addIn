using ExcelDna.Integration;
using ExcelApplicaton = Microsoft.Office.Interop.Excel.Application;

namespace Excel_DNA.Core
{
    public class ExApp
    {
        private static ExcelApplicaton instance;

        public static ExcelApplicaton GetInstance()
        {
            instance ??= (ExcelApplicaton)ExcelDnaUtil.Application;
            return instance;
        }
    }
}
