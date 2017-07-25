using StockAnalysis.excelReader;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI;
using NPOI.HSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace StockAnalysis.excelWriter
{
    class JPNetBuyAnalysis
    {
        private JPNetBuyData data;
        public JPNetBuyAnalysis(JPNetBuyData data)
        {
            this.data = data;
        }
        private IWorkbook openFile(String recordFile)
        {
            IWorkbook workbook;
            using (FileStream file = new FileStream(recordFile, FileMode.Open, FileAccess.ReadWrite))
            {
                workbook = new XSSFWorkbook(file);
            }
            return workbook;
        }
        private void appendData2Record(IWorkbook book, JPNetBuyData data)
        {
            
        }
        private int getLastRow(IWorkbook book)
        {
            ISheet sheet = book.GetSheetAt(0);
            return sheet.LastRowNum;
        }
        
    }
}
