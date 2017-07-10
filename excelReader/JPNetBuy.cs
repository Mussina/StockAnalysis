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
            
            using (FileStream file = new FileStream(dirNetBuy, FileMode.Open, FileAccess.Read))
            {
                workbook = new XSSFWorkbook(dirNetBuy);
            }
            readData();
        }

        private void readData()
        {
            ISheet sheet = workbook.GetSheetAt(0);
            jpNetBuyData = new JPNetBuyData();

            //Dealer
            jpNetBuyData.dealer.orderDiff = sheet.GetRow(Dealer.ROW_DIFF_NET_BUY).GetCell(Dealer.COL_DIFF_NET_BUY).NumericCellValue
                                                                    + sheet.GetRow(Dealer.ROW_DIFF_NET_BUY2).GetCell(Dealer.COL_DIFF_NET_BUY).NumericCellValue;
            //InvestmentTrust          
            jpNetBuyData.dealer.buyIn = sheet.GetRow(InvestmentTrust.ROW_DIFF_NET_BUY).GetCell(InvestmentTrust.COL_DIFF_NET_BUY).NumericCellValue;

            //ForeignInvestor
            jpNetBuyData.foreignInvestor.orderDiff = sheet.GetRow(ForeignInvestor.ROW_DIFF_NET_BUY).GetCell(ForeignInvestor.COL_DIFF_NET_BUY).NumericCellValue;

            // Sum
            jpNetBuyData.totalSum.orderDiff = sheet.GetRow(TotalSum.ROW_DIFF_NET_BUY).GetCell(TotalSum.COL_DIFF_NET_BUY).NumericCellValue;
        }


    }
}
