// Kaggle: C# backend project - Sharanya Radhamohan
// Date: November 20, 2015

IDE used: Visual Studio 2015
Developed on .NET framework 4.5.2
NuGet Package: package id="ExcelDataReader" version="2.1.2.3"
				package id="SharpZipLib" version="0.86.0"
Program.cs
----------
Input to program: Path to filename (example: C:\test\sample.xls)
Output file is created in the same directory as input file. (example: C:\test\sample.csv)

ExcelToCsv: Class has methods to convert a valid .xls or .xlsx file into CSV file.
ConvertExec: In main method, file path is checked if it's valid.
				Instantiates ExcelToCsv and calls the Convert method.

There are no temporary files used. However, an output (.csv) file is created in the file system.