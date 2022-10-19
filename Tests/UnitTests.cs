using Application;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnitTests
{
    [TestClass]
    public class CalculatorTests
    {
        [TestMethod]
        public void EvaluateTestBinaryOperations()
        {
            Assert.AreEqual(Evaluator.GetValue("2+3"), 5);
            Assert.AreEqual(Evaluator.GetValue("7-5"), 2);
            Assert.AreEqual(Evaluator.GetValue("12*12"), 144);
            Assert.AreEqual(Evaluator.GetValue("24/8"), 3);
            Assert.AreEqual(Evaluator.GetValue("17+5-12"), 10);
        }

        [TestMethod]
        public void EvaluateTestUnaryMinus()
        {
            Assert.AreEqual(Evaluator.GetValue("-5"), -5);
            Assert.AreEqual(Evaluator.GetValue("--5"), 5);
            Assert.AreEqual(Evaluator.GetValue("---5"), -5);
            Assert.AreEqual(Evaluator.GetValue("5+-5"), 0);
            Assert.AreEqual(Evaluator.GetValue("5+--5"), 10);
        }

        [TestMethod]
        public void EvaluateTestPow()
        {
            Assert.AreEqual(Evaluator.GetValue("2^3"),8);
            Assert.AreEqual(Evaluator.GetValue("2^2^2"),16);
            Assert.AreEqual(Evaluator.GetValue("2^(2^2)"), 16);
            Assert.AreEqual(Evaluator.GetValue("3^2^(2+2)"), 6561);
        }
        
        [TestMethod]
        public void EvaluateTestIncDec()
        {
            Assert.AreEqual(Evaluator.GetValue("inc(5)"), 6);
            Assert.AreEqual(Evaluator.GetValue("dec(6)"), 5);
            Assert.AreEqual(Evaluator.GetValue("inc(dec(5))"), 5);
            Assert.AreEqual(Evaluator.GetValue("dec(inc(5))"), 5);
        }
        
        [TestMethod]
        public void EvaluateTestMinMax()
        {
            Assert.AreEqual(Evaluator.GetValue("min(3, 3)"), 3);
            Assert.AreEqual(Evaluator.GetValue("min(3, 6)"), 3);
            Assert.AreEqual(Evaluator.GetValue("max(6, 6)"), 6);
            Assert.AreEqual(Evaluator.GetValue("max(3, 6)"), 6);
        }
    }
}
