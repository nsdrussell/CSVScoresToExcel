using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;

namespace ScoresToExcelTests
{
    [TestClass]
    public class ImportUnitTests
    {
        private const string NormalSampleFileName = "NormalSample_export-1582799610.csv";
        private const string BadSampleFileName = "BadSample_export-1582785511.csv";
        private const string GoodSampleFileName = "GoodSample_export-1582775616.csv";
        private const int NumberOfPeopleInNormalSampleFile = 15;

        private string NormalSampleFileNameAndPath => SamplesDirectory + NormalSampleFileName;
        private string BadSampleFileNameAndPath => SamplesDirectory + BadSampleFileName;

        private static string SamplesDirectory
        {
            get
            {
                var directory = new FileInfo(AppDomain.CurrentDomain.BaseDirectory);
                var samplesDirectory = directory.Directory.Parent.Parent.FullName + "\\Samples\\";
                return samplesDirectory;
            }
        }

        [TestMethod]
        public void CheckCanParseGoodFile()
        {
            var parser = new ScoresToExcelApp.CSVParser(NormalSampleFileNameAndPath);
            bool success = parser.CheckCanParse(out _);
            Assert.AreEqual(true, success);
        }

        [TestMethod]
        public void CheckCantParseBadFile()
        {
            var parser = new ScoresToExcelApp.CSVParser(BadSampleFileNameAndPath);
            bool success = parser.CheckCanParse(out _);
            Assert.AreEqual(false, success);
        }

        [TestMethod]
        public void CheckParsesPeopleCorrectlyByCheckingNumberOfItems()
        {
            var parser = new ScoresToExcelApp.CSVParser(NormalSampleFileNameAndPath);
            var result = parser.ParseIntoFileDataset(ScoresToExcelApp.FileDatasetType.CurrentMonth).PeopleWithScores.Count;

            Assert.AreEqual(result, NumberOfPeopleInNormalSampleFile);
        }
    }
}