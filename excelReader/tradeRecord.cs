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
    class TradeRecord
    {
        private IWorkbook tradeRecordbook;
        public TradeRecord(String path)
        {
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                tradeRecordbook = new XSSFWorkbook(file);
            }
        }
    }
}
