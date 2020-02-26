﻿using Microsoft.Win32;
using System;
using System.ComponentModel;
using System.Diagnostics;
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

        public MainWindow()
        {
            InitializeComponent();
            FileNameTextBox.TextChanged += FileNameTextBox_TextChanged;
            if (App.Args != null)
            {
                var fileName = App.Args[0];
                FileNameTextBox.Text = fileName;
            };
        }

        private void FileNameTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!CheckContainsValidFileNameAndPath(FileNameTextBox.Text))
            {
                StatusTextBlock.Text = "The current input is not a valid file path with filename.";
            }
            else
            {
                CSVParser parser = new CSVParser(FileNameTextBox.Text);
                string result;
                if (parser.CheckCanParse(out result))
                {
                    currentDataset = parser.ParseIntoPeopleWithScores();
                    StatusTextBlock.Text = $"File successfully read.";
                    PopulateDataGrid();
                    ExportToExcelButton.IsEnabled = true;
                }
                else
                {
                    StatusTextBlock.Text = "This file can not be read. Are you sure it's the right file?";
                    //ExportToExcelButton.IsEnabled = false;

                    if (!string.IsNullOrEmpty(result))
                        StatusTextBlock.Text += $"{Environment.NewLine}The following error was returned:{Environment.NewLine}{result}";
                }
            }
            if (currentDataset != null)
            {
                StatusTextBlock.Text += $"{Environment.NewLine}Dataset: {currentDataset.ToString()}";
            }
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
                    Header = "Member Average",
                    Binding = new Binding("UnadjustedAverage")
                };
                adjustedAverageColumn = new DataGridTextColumn()
                {
                    Header = "Adjusted Average",
                    Binding = new Binding("AdjustedAverage")
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
            var filename = currentDataset.ExportToExcel();
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

        private void ChooseFileButton_Click(object sender, RoutedEventArgs e)
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
            if (result.HasValue && result.Value) FileNameTextBox.Text = filepicker.FileName;
        }
    }
}