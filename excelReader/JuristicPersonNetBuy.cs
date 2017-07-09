using NPOI.HSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StockAnalysis.excelReader
{
    class JuristicPersonNetBuy
    {
        private HSSFWorkbook workbook;
        public JuristicPersonNetBuy(String dailyNetBuy)
        {
            
            using (FileStream file = new FileStream(AppDomain.CurrentDomain.BaseDirectory + "sample.xls", FileMode.Open, FileAccess.Read))
            {
                workbook = new HSSFWorkbook(file);
            }
        }




    }
}
