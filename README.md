# CSVScoresToExcel
This is a small application that generates spreadsheets for a sports club. The outputs are intended to be viewed by members to see how they are doing. 

There is a system in place already to keep track of scores, which includes the ability to export scores to .csv format. Currently, heads of sections must do a manual process of data cleansing and copying into spreadsheets which is time consuming, and as a result often delayed because time has not been found to do it.

## Download
Latest version of the application can be found here.

https://github.com/nsdrussell/CSVScoresToExcel/releases/

## Future
At present the application matches the original specification, so I am unlikely to do much more except bug fixing when they are found. However, if you'd like this adapted for your own club or own purposes, either let me know by submitting an issue with the [issue tracker](https://github.com/nsdrussell/CSVScoresToExcel/issues) and I'll see if I can help, or make a new branch and do so yourself.

I might decide to make an installer for the application at some point.

## Passing files as arguments
The filename can be passed as an argument to read it on opening. This will save having to select the current month file manually. There is more than one way to do this.
* Easiest: Drag the .csv file onto CSVScoresToExcelApp.exe (the program itself) and let go.
* Permanent: Right click in Windows Explorer on your desired .csv file->Open With...-> Tick the "Always use this app to open .csv files" tickbox> Choose another app-> Choose the application executable. Now when you double click a csv it will always open with this application.
* Technical: Can do it from the command line; like `ScoresToExcelApp "C:\Files\NormalSample_export-1582799610.csv"`.
## Build instructions
In Visual Studio, load the .sln file under `/src/ScoresToExcelApp/`, then build and run. You will need WPF (Windows Presentation Foundation) and .NET Framework 4.7.2 enabled. 

I wrote and built this using VS2019.

## Information
### Input
The expected input looks like this.

`sportname_export-unixtimeofexport.csv`
```
Name,Member Nos,Average,Scores
"person a",1234,92.5,92,93
"person b",5678,81.67,81,82,82
```
The number of columns is variable depending on the number of scores that have been put in.

Each row is separated by a line feed (\n).

I have included samples under `/src/ScoresToExcelApp/Samples/`. GitHub reports that there are errors with the files, but they are genuinely as they should be.
### Data handling
I was tempted to try parsing the entire file to a DataTable, but unfortunately there is a varying number of columns and this method would not be up to scratch as a result. I take each line from the file after the header row and parse them into PersonWithScore objects.

The application does a small amount of data cleansing, - scores that are too different from a trimmed mean are kept separately so they can be marked on output and an "adjusted mean" can be calculated that ignores them.

Data cleansing is done so that if a member ordinarily scores in the 70s and inputs a score in the 90s by mistake, it can be ignored. And likewise if they ordinarily score in the 90s and put in a 70 by mistake it can be ignored.

## Further information
### Arithmetic
I adapted the TruncatedMean method from the [Accord.NET framework](https://github.com/accord-net/framework). I don't take the truncated mean as the average, but rather use it to remove erroneous scores from the scores used then calculate an average.
### Exporting to Excel format
I chose to use the [EPPlus framework](https://github.com/JanKallman/EPPlus) to output the data to excel as I am familiar with it and it works well.
