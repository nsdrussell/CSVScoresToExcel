# CSVScoresToExcel
This is a small application that generates spreadsheets for a sports club. The outputs are intended to be viewed by members to see how they are doing. 

There is a system in place already to keep track of scores, which includes the ability to scores to .csv format. Currently, heads of sections must do a manual process of data cleaning and copying into spreadsheets which is time consuming, and as a result often delayed because time has not been found to do it.

Hopefully this should speed things up.

## Other information
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
### Data
The application does a small amount of data cleaning, scores that are too different from a trimmed mean are marked as such.
