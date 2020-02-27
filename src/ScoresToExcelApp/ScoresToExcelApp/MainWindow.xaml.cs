using Microsoft.Win32;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace ScoresToExcelApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private FileDataset currentDataset;
        private FileDataset previousDataset;

        public MainWindow()
        {
            InitializeComponent();
            CurrentFileNameTextBox.TextChanged += CurrentFileNameTextBox_TextChanged;
            if (App.Args != null)
            {
                var fileName = App.Args[0];
                CurrentFileNameTextBox.Text = fileName;
            };
        }

        private void PreviousFileNameTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!CheckContainsValidFileNameAndPath(CurrentFileNameTextBox.Text))
            {
                StatusTextBlock.Text = "The previous input is not a valid file path with filename.";
            }
            else
            {
                CSVParser parser = new CSVParser(PreviousFileNameTextBox.Text);
                if (parser.CheckCanParse(out string result))
                {
                    previousDataset = parser.ParseIntoFileDataset(FileDatasetType.PreviousMonth);
                    SetCurrentDatasetPreviousAverages();
                    StatusTextBlock.Text = $"File successfully read.";
                    PopulateDataGrid();
                    DateCalendar.SelectedDateChanged += DateCalendar_SelectedDateChanged;
                    ExportToExcelButton.IsEnabled = true;
                }
                else
                {
                    StatusTextBlock.Text = "File can't be read.";

                    if (!string.IsNullOrEmpty(result))
                        StatusTextBlock.Text += $" Error:{Environment.NewLine}{result}";
                }
            }
            if (previousDataset != null)
            {
                StatusTextBlock.Text += $"{Environment.NewLine}Last month dataset: {previousDataset}";
            }
        }

        private void SetCurrentDatasetPreviousAverages()
        {
            if (currentDataset != null && previousDataset != null)
            {
                foreach (var memberInCurrentDataset in currentDataset.PeopleWithScores)
                {
                    if (previousDataset.PeopleWithScores.Any(member => member.MemberName == memberInCurrentDataset.MemberName))
                    {
                        var memberInPreviousDataset = previousDataset.PeopleWithScores.First(member => member.MemberName == memberInCurrentDataset.MemberName);

                        memberInCurrentDataset.SetPreviousAverage(memberInPreviousDataset.AdjustedAverage);
                    }
                    else //if not reset to zero.
                    {
                        memberInCurrentDataset.SetPreviousAverage(null);
                    }
                }
                ScoresDataGrid.Items.Refresh();
            }
        }

        private void CurrentFileNameTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!CheckContainsValidFileNameAndPath(CurrentFileNameTextBox.Text))
            {
                StatusTextBlock.Text = "The current input is not a valid file path with filename.";
            }
            else
            {
                CSVParser parser = new CSVParser(CurrentFileNameTextBox.Text);
                string result;
                if (parser.CheckCanParse(out result))
                {
                    //if have set the start date and end date
                    if (DateCalendar.SelectedDate != null)
                    {
                        currentDataset = parser.ParseIntoFileDataset(FileDatasetType.CurrentMonth, (DateTime)DateCalendar.SelectedDate);
                    }
                    else
                    {
                        currentDataset = parser.ParseIntoFileDataset(FileDatasetType.CurrentMonth);

                        if (DateCalendar.SelectedDate == null)
                            DateCalendar.SelectedDate = currentDataset.DateOfScores;
                        else
                            currentDataset.DateOfScores = (DateTime)DateCalendar.SelectedDate;
                    }
                    StatusTextBlock.Text = $"File successfully read.";
                    PopulateDataGrid();

                    DateCalendar.SelectedDateChanged += DateCalendar_SelectedDateChanged;
                    PreviousFileNameTextBox.TextChanged += PreviousFileNameTextBox_TextChanged;

                    PreviousChooseFileButton.IsEnabled = true;
                    PreviousFileNameTextBox.IsEnabled = true;
                    ExportToExcelButton.IsEnabled = true;
                }
                else
                {
                    StatusTextBlock.Text = "File can't be read.";
                    //ExportToExcelButton.IsEnabled = false;

                    if (!string.IsNullOrEmpty(result))
                        StatusTextBlock.Text += $" Error:{Environment.NewLine}{result}";
                }
            }
            if (currentDataset != null)
            {
                StatusTextBlock.Text += $"{Environment.NewLine}Dataset: {currentDataset.ToString()}";
                SetCurrentDatasetPreviousAverages();
            }
        }

        private void DateCalendar_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            currentDataset.DateOfScores = (DateTime)(sender as DatePicker).SelectedDate;
        }

        private bool CheckContainsValidFileNameAndPath(string text)
        {
            return text.EndsWith(".csv") && text.Contains('\\');
        }

        private void PopulateDataGrid()
        {
            DataGridColumn adjustedAverageColumn;
            if (!ScoresDataGrid.Columns.Any())
            {
                DataGridTextColumn nameCol = new DataGridTextColumn()
                {
                    Header = "Member Name",
                    Binding = new Binding("MemberName")
                };
                adjustedAverageColumn = new DataGridTextColumn()
                {
                    Header = "Current",
                    Binding = new Binding("AdjustedAverage") { StringFormat = "N2" }
                };
                DataGridTextColumn previousAverageColumn = new DataGridTextColumn()
                {
                    Header = "Previous",
                    Binding = new Binding("PreviousAverage") { StringFormat = "N2" }
                };
                DataGridTextColumn scoreColumn = new DataGridTextColumn()
                {
                    Header = "Member Scores",
                    Binding = new Binding("ScoresList")
                };
                DataGridTextColumn badScoresColumn = new DataGridTextColumn()
                {
                    Header = "Erroneous Scores",
                    Binding = new Binding("ErroneousScoresList")
                };
                DataGridTextColumn categoryColumn = new DataGridTextColumn()
                {
                    Header = "Category",
                    Binding = new Binding("ReadableCategory")
                };

                ScoresDataGrid.Columns.Add(nameCol);
                ScoresDataGrid.Columns.Add(previousAverageColumn);
                ScoresDataGrid.Columns.Add(adjustedAverageColumn);
                ScoresDataGrid.Columns.Add(scoreColumn);
                ScoresDataGrid.Columns.Add(badScoresColumn);
                ScoresDataGrid.Columns.Add(categoryColumn);
            }
            else
            {
                adjustedAverageColumn = ScoresDataGrid.Columns[2];
                ScoresDataGrid.Items.Clear();
            }

            currentDataset.PeopleWithScores.ForEach(member => { ScoresDataGrid.Items.Add(member); });

            SortDataGrid(ScoresDataGrid, adjustedAverageColumn);
        }

        /// <summary>
        /// Sort the datagridview based on the column. Method found here: https://stackoverflow.com/a/19952233
        /// </summary>
        public static void SortDataGrid(DataGrid dataGrid, DataGridColumn column)
        {
            // Clear current sort descriptions
            dataGrid.Items.SortDescriptions.Clear();

            // Add the new sort description
            dataGrid.Items.SortDescriptions.Add(new SortDescription(column.SortMemberPath, ListSortDirection.Descending));

            // Apply sort
            foreach (var col in dataGrid.Columns) col.SortDirection = null;

            column.SortDirection = ListSortDirection.Descending;

            // Refresh items to display sort
            dataGrid.Items.Refresh();
        }

        private void ExportToExcelButton_Click(object sender, RoutedEventArgs e)
        {
            string filename;
            try
            {
                filename = currentDataset.ExportToExcel();
            }
            catch (IOException)
            {
                MessageBox.Show("For some reason the file can't be saved. Is there currently one you just made open? If there is, shut it.", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            var messageBoxResult = MessageBox.Show($"Output successful. Saved to:{Environment.NewLine}{filename}{Environment.NewLine}" +
                $"Would you like to open the export?",
                "Success", MessageBoxButton.YesNo, MessageBoxImage.Information);

            if (messageBoxResult == MessageBoxResult.Yes)
            {
                Process.Start(filename);
                if ((bool)CloseCheckBox.IsChecked)
                {
                    Close();
                }
            }
        }

        private void CurrentChooseFileButton_Click(object sender, RoutedEventArgs e)
        {
            var myDocuments = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var filepicker = new OpenFileDialog()
            {
                InitialDirectory = myDocuments,
                Multiselect = false,
                Filter = "Flat file database (*.csv)|*.csv",
                DefaultExt = " *.csv"
            };

            var result = filepicker.ShowDialog();
            if (result.HasValue && result.Value) CurrentFileNameTextBox.Text = filepicker.FileName;
        }

        private void PreviousChooseFileButton_Click(object sender, RoutedEventArgs e)
        {
            var myDocuments = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var filepicker = new OpenFileDialog()
            {
                InitialDirectory = myDocuments,
                Multiselect = false,
                Filter = "Flat file database (*.csv)|*.csv",
                DefaultExt = " *.csv"
            };

            var result = filepicker.ShowDialog();
            if (result.HasValue && result.Value) PreviousFileNameTextBox.Text = filepicker.FileName;
        }

        private void SourceLabel_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            Process.Start("https://github.com/nsdrussell/CSVScoresToExcel");
        }
    }
}