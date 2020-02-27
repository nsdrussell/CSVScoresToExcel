using Microsoft.VisualStudio.TestTools.UnitTesting;
using ScoresToExcelApp;
using System.Linq;

namespace ScoresToExcelTests
{
    [TestClass]
    public class ArithmeticUnitTests
    {
        private string[] normalScores = { "90", "91", "92", "91", "91" };
        private string[] abnormalScores = { "90", "91", "92", "91", "91", "20", "45" };
        private string[] erroneousScores = { "20", "45" };

        [TestMethod]
        public void TestTruncatedAverageDoesntTruncateWantedScores()
        {
            var result = ArithmaticHelpers.TruncatedMean(normalScores);
            var expectedResult = 91d;
            Assert.AreEqual(result, expectedResult);
        }

        [TestMethod]
        public void TestTruncatedAverageTruncatesUnWantedScores()
        {
            var result = ArithmaticHelpers.TruncatedMean(abnormalScores);
            var expectedResult = abnormalScores.Select(score => int.Parse(score)).Average();
            Assert.AreNotEqual(result, expectedResult);
        }

        [TestMethod]
        public void TestAblityToDetectErroneousScoresFromTruncatedMean()
        {
            var truncatedMean = ArithmaticHelpers.TruncatedMean(abnormalScores);
            var result = ArithmaticHelpers.GetErroneousScoresGivenTruncatedMean(abnormalScores.Select(score => int.Parse(score)).ToArray(), truncatedMean);

            var expectedResult = erroneousScores.Select(score => int.Parse(score)).ToArray();

            Assert.AreNotEqual(result, expectedResult);
        }
    }
}