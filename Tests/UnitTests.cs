using Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnitTests
{
    [TestClass()]
    public class CalculatorTests
    {
        [TestMethod()]
        public void EvaluateTestUnaryMinus()
        {
            Assert.AreEqual(Calculator.Evaluate("-5"), -5);
            Assert.AreEqual(Calculator.Evaluate("--5"), 5);
            Assert.AreEqual(Calculator.Evaluate("---5"), -5);
            Assert.AreEqual(Calculator.Evaluate("5+-5"), 0);
            Assert.AreEqual(Calculator.Evaluate("5+--5"), 10);
        }

        [TestMethod()]
        public void EvaluateTestPow()
        {
            Assert.AreEqual(Calculator.Evaluate("2^3"),8);
            Assert.AreEqual(Calculator.Evaluate("2^2^2"),16);
            Assert.AreEqual(Calculator.Evaluate("2^(2^2)"), 16);
            Assert.AreEqual(Calculator.Evaluate("3^2^(2+2)"), 6561);
        }
        [TestMethod()]
        public void EvaluateTestIncDec()
        {
            Assert.AreEqual(Calculator.Evaluate("inc(5)"), 6);
            Assert.AreEqual(Calculator.Evaluate("dec(6)"), 5);
            Assert.AreEqual(Calculator.Evaluate("inc(dec(5))"), 5);
            Assert.AreEqual(Calculator.Evaluate("dec(inc(5))"), 5);
        }
    }
}
