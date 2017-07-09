using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI;
using NPOI.HSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StockAnalysis.excelReader
{
    class DailyNetBuyIndex
    {
      //  public const int 
    }
   public class JPNetBuy
    {
        // net buy excel
        private IWorkbook workbook;

        private JPNetBuyData jpNetBuyData;

        public JPNetBuy(String dirNetBuy)
        {
            jpNetBuyData = new JPNetBuyData();
            using (FileStream file = new FileStream(dirNetBuy, FileMode.Open, FileAccess.Read))
            {
                workbook = new XSSFWorkbook(dirNetBuy);
            }
            readData();
        }

        private void readData()
        {
            ISheet sheet = workbook.GetSheetAt(0);
            Double a = sheet.GetRow(2).GetCell(3).NumericCellValue;
        }


    }
}
