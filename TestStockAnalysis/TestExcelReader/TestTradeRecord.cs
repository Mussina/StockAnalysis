using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
//using StockAnalysis.excelReader;

namespace StockAnalysis.excelReader
{
    [TestClass]
    public class TestTradeRecord
    {
        [TestMethod]
        public void TestGetDate()
        {
            Assert.AreEqual<String>( "2017/7/12", TradeRecord.getDate());
        }

        [TestMethod]
        public void TestDiffDate()
        {
            Assert.AreEqual<int>(3 , TradeRecord.getDiffDate("2017/7/10","2017/7/13"));
        }

    }
}
