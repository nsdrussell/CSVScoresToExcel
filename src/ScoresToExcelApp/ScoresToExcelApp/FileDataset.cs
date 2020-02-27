using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;

namespace ScoresToExcelApp
{
    internal class FileDataset
    {
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }

        private const int ExcelCategoryNameColumnIndex = 1;
        private const int ExcelNameColumnIndex = 2;
        private const int ExcelAverageColumnIndex = 3;
        private const int ExcelPreviousAverageColumnIndex = 4;
        private const int ExcelDifferenceColumnIndex = 5;
        private const int ExcelScoresMinimumColumnIndex = 6;

        public List<PersonWithScores> PeopleWithScores { get; }
        public FileDatasetType DatasetType { get; }
        public string SportName { get; }
        public DateTime ExportDateTime { get; }

        public FileDataset(List<PersonWithScores> peopleWithScores, string fullFileName, FileDatasetType datasetType)
        {
            this.PeopleWithScores = peopleWithScores;
            this.DatasetType = datasetType;
            var fileName = Path.GetFileNameWithoutExtension(fullFileName);

            //filename format is sportname_export-1234567890.csv so can derive sportname and export date from it
            SportName = fileName.Split('_')[0];

            long epochSeconds = long.Parse(fileName.Split('-')[1]);

            ExportDateTime = new DateTime(1970, 1, 1, 0, 0, 0, 0, DateTimeKind.Unspecified).AddSeconds(epochSeconds);
        }

        public void SetPreviousAverage(string name, double average)
        {
            if (PeopleWithScores.Any(person => person.MemberName == name))
            {
                var memberWithPreviousScore = PeopleWithScores.First(person => person.MemberName == name);
                memberWithPreviousScore.SetPreviousAverage(average);
            }
        }

        public DataTable GetDataTableOfScores()
        {
            var results = new DataTable();
            results.Columns.Add("", typeof(string));
            return results;
        }

        /// <summary>
        /// Returns as string like "SPORTNAME Scores yyyyMMdd
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return $"{SportName}, Export date: {ExportDateTime.ToString("yyyyMMdd")}";
        }

        public FileInfo GetNewFileFileInfo()
        {
            var myDocuments = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var newFileName = $"{SportName} Scores {StartDate.ToString("yyyyMMdd")} - {EndDate.ToString("yyyyMMdd")}.xlsx";

            FileInfo fileInfo = new FileInfo(myDocuments + "\\" + newFileName);
            return fileInfo;
        }

        /// <summary>
        /// Export to excel. Returns the filename.
        /// </summary>
        /// <returns></returns>
        internal string ExportToExcel()
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                package.Workbook.Worksheets.Add(SportName + " Scores");
                var sheet = package.Workbook.Worksheets[1];

                int minimumRowIndex = 1;
                //headers
                sheet.Cells[minimumRowIndex, ExcelNameColumnIndex].Value = "Name";
                sheet.Cells[minimumRowIndex, ExcelAverageColumnIndex].Value = "Adjusted Average";
                sheet.Cells[minimumRowIndex, ExcelPreviousAverageColumnIndex].Value = "Previous Average";
                sheet.Cells[minimumRowIndex, ExcelDifferenceColumnIndex].Value = "Difference";
                sheet.Cells[minimumRowIndex, ExcelScoresMinimumColumnIndex].Value = "Scores";

                //set bold
                sheet.Column(ExcelDifferenceColumnIndex).Style.Font.Bold = true;
                sheet.Row(minimumRowIndex++).Style.Font.Bold = true;

                //each category
                var categories = PeopleWithScores.Select(s => s.Category).Distinct().OrderBy(categoryName => categoryName);
                foreach (var category in categories)
                {
                    //Category
                    sheet.Cells[minimumRowIndex, ExcelCategoryNameColumnIndex].Value =
                        PeopleWithScores.First(member => member.Category == category).ReadableCategory;

                    sheet.Row(minimumRowIndex++).Style.Font.Bold = true;

                    var membersInCategory = PeopleWithScores
                        .Where(member => member.Category == category)
                        .OrderByDescending(member => member.AdjustedAverage);

                    foreach (var person in membersInCategory)
                    {
                        sheet.Cells[minimumRowIndex, ExcelNameColumnIndex].Value = person.MemberName;
                        sheet.Cells[minimumRowIndex, ExcelAverageColumnIndex].Value = person.AdjustedAverage;
                        sheet.Cells[minimumRowIndex, ExcelPreviousAverageColumnIndex].Value = person.PreviousAverage;
                        sheet.Cells[minimumRowIndex, ExcelDifferenceColumnIndex].Formula = $"=C{minimumRowIndex}-D{minimumRowIndex}";
                        for (int i = 0; i < person.Scores.Length; i++)
                        {
                            sheet.Cells[minimumRowIndex, ExcelScoresMinimumColumnIndex + i].Value
                                = person.Scores[i];

                            if (person.ErroneousScores.Contains(person.Scores[i]))
                            {
                                sheet.Cells[minimumRowIndex, ExcelScoresMinimumColumnIndex + i]
                                    .Style.Font.Bold = true;
                            }
                        }
                        minimumRowIndex++;
                    }
                }
                sheet.Cells.AutoFitColumns();
                sheet.Column(ExcelAverageColumnIndex).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                sheet.Column(ExcelAverageColumnIndex).Style.Numberformat.Format =
                sheet.Column(ExcelPreviousAverageColumnIndex).Style.Numberformat.Format =
                sheet.Column(ExcelDifferenceColumnIndex).Style.Numberformat.Format = "0.00";

                for (int i = ExcelScoresMinimumColumnIndex; i <= sheet.Dimension.End.Column; i++) { sheet.Column(i).Width = 4; }

                //rule for difference in averages
                ExcelAddress cells = new ExcelAddress(1, ExcelDifferenceColumnIndex, sheet.Dimension.End.Row, ExcelDifferenceColumnIndex);
                var cfRule = sheet.ConditionalFormatting.AddThreeColorScale(cells);
                cfRule.HighValue.Type =
                cfRule.MiddleValue.Type =
                cfRule.LowValue.Type = eExcelConditionalFormattingValueObjectType.Num;

                cfRule.HighValue.Color = Color.FromArgb(198, 239, 206);
                cfRule.MiddleValue.Color = Color.White;
                cfRule.LowValue.Color = Color.FromArgb(255, 199, 206);

                cfRule.HighValue.Value = 7.5;
                cfRule.MiddleValue.Value = 0;
                cfRule.LowValue.Value = -10;

                var fileInfo = GetNewFileFileInfo();
                package.SaveAs(fileInfo);
                return fileInfo.FullName;
            }
        }
    }
}