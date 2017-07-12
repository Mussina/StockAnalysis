using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StockAnalysis.excelReader
{
     public class TradeRecord
    {
        private IWorkbook tradeRecordbook;
        public TradeRecord(String path)
        {
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                tradeRecordbook = new XSSFWorkbook(file);
            }
        }

        public static String getDate()
        {
            DateTime dt = DateTime.Now;
            return dt.ToShortDateString();
        }
        public static int getDiffDate(String earlyDay, String laterDay)
        {
            DateTime dt1 = Convert.ToDateTime(earlyDay);
            DateTime dt2 = Convert.ToDateTime(laterDay);
            TimeSpan span = dt2.Subtract(dt1);
            return span.Days ;
        }

    }
}
