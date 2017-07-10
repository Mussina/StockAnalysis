using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using StockAnalysis.excelReader;

namespace TestStockAnalysis
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            JPNetBuy jpNetBuy = new JPNetBuy("D:\\Coding4Fun\\3MainJuristicPersonRecord\\20170707.xlsx");

        }
    }
}
