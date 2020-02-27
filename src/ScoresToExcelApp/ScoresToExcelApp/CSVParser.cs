using System;
using System.Collections.Generic;
using System.IO;

namespace ScoresToExcelApp
{
    public class CSVParser
    {
        public string FileName { get; }

        public CSVParser(string fileName)
        {
            FileName = fileName;
        }

        public FileDataset ParseIntoFileDataset(FileDatasetType datasetType)
        {
            var peopleWithScores = new List<PersonWithScores>();
            using (var reader = new StreamReader(FileName))
            {
                reader.ReadLine(); //ignore the first line as is just headers.

                //now read the rest of the file to the end and add people based on that.
                while (!reader.EndOfStream)
                {
                    var row = reader.ReadLine().Trim();
                    if (string.IsNullOrEmpty(row)) continue;

                    peopleWithScores.Add(new PersonWithScores(row));
                }
            }

            return new FileDataset(peopleWithScores, FileName, datasetType);
        }

        public FileDataset ParseIntoFileDataset(FileDatasetType datasetType, DateTime startDate, DateTime endDate)
        {
            var peopleWithScores = new List<PersonWithScores>();
            using (var reader = new StreamReader(FileName))
            {
                reader.ReadLine(); //ignore the first line as is just headers.

                //now read the rest of the file to the end and add people based on that.
                while (!reader.EndOfStream)
                {
                    var row = reader.ReadLine().Trim();
                    if (string.IsNullOrEmpty(row)) continue;

                    peopleWithScores.Add(new PersonWithScores(row));
                }
            }

            return new FileDataset(peopleWithScores, FileName, datasetType, startDate, endDate);
        }

        /// <summary>
        /// Checks if can parse the file.
        /// If can parse, result = string.Empty, otherwise the error message if can't.
        /// </summary>
        public bool CheckCanParse(out string result)
        {
            string firstLine = result = string.Empty;
            try
            {
                using (StreamReader reader = new StreamReader(FileName))
                {
                    firstLine = reader.ReadLine();
                }
            }
            catch (Exception ex)
            {
                result = ex.Message;
            }

            return firstLine == Properties.Resources.FileLineHeaderText;
        }
    }
}