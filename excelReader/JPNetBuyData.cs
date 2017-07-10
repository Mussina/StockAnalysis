using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StockAnalysis.excelReader
{
    /// <summary>
    /// JuristicPersonNetBuy's data
    /// </summary>
    class JPNetBuyData
    {
        public const int ROW_DIFF_NET_BUY = 6;
        public const int COL_DIFF_NET_BUY = 3;
        public Dealer dealer = new Dealer();
        public InvestmentTrust investmentTrust = new InvestmentTrust();
        public ForeignInvestor foreignInvestor = new ForeignInvestor();
        public TotalSum totalSum = new TotalSum();
    }

    class JuristicPersonData
    {
        public Double buyIn;
        public Double sellOut;
        public Double orderDiff;
    }
    class Dealer : JuristicPersonData
    {
        public const int ROW_DIFF_NET_BUY = 2;
        public const int ROW_DIFF_NET_BUY2 = 3;
        public const int COL_DIFF_NET_BUY = 3;
    }
    class InvestmentTrust : JuristicPersonData
    {
        public const int ROW_DIFF_NET_BUY = 4;
        public const int COL_DIFF_NET_BUY = 3;
    }
    class ForeignInvestor : JuristicPersonData
    {
        public const int ROW_DIFF_NET_BUY = 5;
        public const int COL_DIFF_NET_BUY = 3;
    }

    class TotalSum : JuristicPersonData
    {
        public const int ROW_DIFF_NET_BUY = 6;
        public const int COL_DIFF_NET_BUY = 3;
    }


}
