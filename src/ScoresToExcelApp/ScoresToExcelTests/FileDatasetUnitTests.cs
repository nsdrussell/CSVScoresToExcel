using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ScoresToExcelTests
{
    [TestClass]
    public class FileDatasetUnitTests
    {
        private const string NormalSampleFileName = "NormalSample_export-1582799610.csv";
        private const string BadSampleFileName = "BadSample_export-1582785511.csv";
        private const string GoodSampleFileName = "GoodSample_export-1582775616.csv";
        private const int NumberOfPeopleInNormalSampleFile = 15;

        private string NormalSampleFileNameAndPath => SamplesDirectory + NormalSampleFileName;
        private string BadSampleFileNameAndPath => SamplesDirectory + BadSampleFileName;

        string[] membersInNormalSample = { "Member A","Member B","Member C","Member D","Member E","Member F",
"Member G","Member H","Member I","Member J","Member K","Member L",
"Member M","Member N","Member O",};

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
        public void CheckNamesSetCorrectly()
        {
            var parser = new ScoresToExcelApp.CSVParser(NormalSampleFileNameAndPath);
            var results = parser.ParseIntoFileDataset(ScoresToExcelApp.FileDatasetType.CurrentMonth).PeopleWithScores.Select(person => person.MemberName).ToArray();

            //check contain eachother
            foreach (var item in results)
            {
                if (!membersInNormalSample.Contains(item)) Assert.Fail("a result was not in results that should have been.");
            }
            foreach (var item in membersInNormalSample)
            {
                if (!results.Contains(item)) Assert.Fail("a result was not in results that should have been.");
            }
            Assert.IsTrue(true);
        }
    }
}