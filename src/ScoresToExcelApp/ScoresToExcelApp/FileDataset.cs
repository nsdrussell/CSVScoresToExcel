using OfficeOpenXml;
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
        private const int ExcelCategoryNameColumnIndex = 1;
        private const int ExcelNameColumnIndex = 2;
        private const int ExcelAverageColumnIndex = 3;
        private const int ExcelScoresMinimumColumnIndex = 4;

        public List<PersonWithScores> PeopleWithScores { get; }
        public string SportName { get; }
        public DateTime ExportDateTime { get; }

        public FileDataset(List<PersonWithScores> peopleWithScores, string fullFileName)
        {
            this.PeopleWithScores = peopleWithScores;

            var fileName = fullFileName.Split('\\').LastOrDefault().TrimEnd('v', 's', 'c', '.');
            //filename format is sportname_export-1234567890.csv so ca derive sportname and date of export from it
            SportName = fileName.Split('_')[0];

            long epochSeconds = long.Parse(fileName.Split('-')[1]);

            ExportDateTime = new DateTime(1970, 1, 1, 0, 0, 0, 0, DateTimeKind.Unspecified).AddSeconds(epochSeconds);
        }

        public DataTable GetDataTableOfScores()
        {
            var results = new DataTable();
            results.Columns.Add("", typeof(string));
            return results;
        }

        public override string ToString()
        {
            return $"{SportName}, Export date: {ExportDateTime.ToString("dd/MM/yyyy")}";
        }

        internal void ExportToExcel()
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                package.Workbook.Worksheets.Add(SportName + " Scores");
                var sheet = package.Workbook.Worksheets[1];

                int minimumRowIndex = 1;
                //headers
                sheet.Cells[minimumRowIndex, ExcelNameColumnIndex].Value = "Name";
                sheet.Cells[minimumRowIndex, ExcelAverageColumnIndex].Value = "Adjusted Average";
                sheet.Cells[minimumRowIndex, ExcelScoresMinimumColumnIndex].Value = "Scores";
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
                        for (int i = 0; i < person.Scores.Length; i++)
                        {
                            sheet.Cells[minimumRowIndex, ExcelScoresMinimumColumnIndex + i].Value
                                = person.Scores[i];

                            if (person.ErroneousScores.Contains(person.Scores[i]))
                            {
                                sheet.Cells[minimumRowIndex, ExcelScoresMinimumColumnIndex + i]
                                    .Style.Font.Bold = true;
                                sheet.Cells[minimumRowIndex, ExcelScoresMinimumColumnIndex + i]
                                    .Style.Fill.BackgroundColor.SetColor(Color.Red);
                            }
                        }
                        minimumRowIndex++;
                    }
                }
                sheet.Cells.AutoFitColumns();
                sheet.Column(ExcelAverageColumnIndex).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                for (int i = ExcelScoresMinimumColumnIndex; i <= sheet.Dimension.End.Column; i++) { sheet.Column(i).Width = 3; }

                var myDocuments = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                var newFileName = $"{SportName} {DateTime.Now.ToString("yyyyMMdd")}.xlsx";
                FileInfo fi = new FileInfo(myDocuments + "\\" + newFileName);
                package.SaveAs(fi);
            }
        }
    }
}