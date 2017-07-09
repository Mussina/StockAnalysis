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
        public const int rowDiffNetBuy = 6;
        public const int colDiffNetBuy = 3;
        public Dealer dealer;
        public InvestmentTrust investmentTrust;
        public ForeignInvestor foreignInvestor;
    }

    class JuristicPersonData
    {
        public long buyIn;
        public long sellOut;
        public long orderDiff;
    }
    class Dealer : JuristicPersonData
    {
        public const int rowDiffNetBuy = 2;
        public const int colDiffNetBuy = 3;
    }
    class InvestmentTrust : JuristicPersonData
    {
        public const int rowDiffNetBuy = 4;
        public const int colDiffNetBuy = 3;
    }
    class ForeignInvestor : JuristicPersonData
    {
        public const int rowDiffNetBuy = 5;
        public const int colDiffNetBuy = 3;
    }
   
    
}
