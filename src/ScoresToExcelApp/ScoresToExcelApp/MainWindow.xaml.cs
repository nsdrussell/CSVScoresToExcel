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
            PreviousFileNameTextBox.TextChanged += PreviousFileNameTextBox_TextChanged;
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
                    previousDataset = parser.ParseIntoPeopleWithScores(FileDatasetType.PreviousMonth);
                    StatusTextBlock.Text = $"File successfully read.";
                    PopulateDataGrid();
                    StartDateCalendar.SelectedDateChanged += StartDateCalendar_SelectedDateChanged;
                    EndDateCalendar.SelectedDateChanged += EndDateCalendar_SelectedDateChanged;
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
                StatusTextBlock.Text += $"{Environment.NewLine}Last month dataset: {previousDataset.ToString()}";
                SetCurrentDatasetPreviousAverages();
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
                    currentDataset = parser.ParseIntoPeopleWithScores(FileDatasetType.CurrentMonth);
                    StatusTextBlock.Text = $"File successfully read.";
                    PopulateDataGrid();
                    StartDateCalendar.SelectedDateChanged += StartDateCalendar_SelectedDateChanged;
                    EndDateCalendar.SelectedDateChanged += EndDateCalendar_SelectedDateChanged;
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

        private void EndDateCalendar_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            currentDataset.EndDate = (DateTime)(sender as DatePicker).SelectedDate;
        }

        private void StartDateCalendar_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            currentDataset.StartDate = (DateTime)(sender as DatePicker).SelectedDate;
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
                DataGridTextColumn averageColumn = new DataGridTextColumn()
                {
                    Header = "Average",
                    Binding = new Binding("UnadjustedAverage") { StringFormat = "N2" }
                };
                adjustedAverageColumn = new DataGridTextColumn()
                {
                    Header = "Adjusted",
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
                ScoresDataGrid.Columns.Add(averageColumn);
                ScoresDataGrid.Columns.Add(adjustedAverageColumn);
                ScoresDataGrid.Columns.Add(previousAverageColumn);
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
        /// Sort the datagridview based on the column. Found here: https://stackoverflow.com/a/19952233
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