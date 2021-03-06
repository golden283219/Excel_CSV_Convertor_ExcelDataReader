using System;
using System.IO;
using System.Data;
using Excel;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Configuration;
using Microsoft.VisualBasic.FileIO;

namespace ExcelToCsvApp
{
    class ExcelToCsv
    {
        FileStream inputStream;
        IExcelDataReader excelReader;
        StreamWriter writer;
        DataSet result;

        public ExcelToCsv(string ipfilename)
        {
            try
            {
                inputStream = File.Open(ipfilename, FileMode.Open, FileAccess.Read);
                // Read from a *.xls file (97-2003 format)
                if (Path.GetExtension(ipfilename) == ".xls")
                    excelReader = ExcelReaderFactory.CreateBinaryReader(inputStream);
                // Read from a *.xlsx file (2007 format)
                else
                    excelReader = ExcelReaderFactory.CreateOpenXmlReader(inputStream);
                // DataSet - The result of each spreadsheet will be created in the result.Tables
                result = excelReader.AsDataSet();
            }
            catch (Exception e)
            {
                Console.WriteLine("\nAn exception occured while trying to read the input file.");
                Console.WriteLine(e.ToString());
            }
        }
       ~ExcelToCsv() //Destructor to close file streams
        {
            if (excelReader != null)
                excelReader.Close();
            if (inputStream != null)
                inputStream.Close();
        }
        public void Convert(string opfilename)
        {
            // excelReader.IsFirstRowAsColumnNames = true;
            writer = new StreamWriter(opfilename);
            string s = "";
            foreach (DataTable table in result.Tables)
            {
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    s = "";
                    for (int j = 0; j < table.Columns.Count; j++)
                    {
                        writer.AutoFlush = true;
                        //Console.WriteLine("\"" + table.Rows[i].ItemArray[j] + "\";");
                        s += table.Rows[i].ItemArray[j] + ",";
                    }
                    s = s.Substring(0, s.Length - 1);
                    //Console.WriteLine(s);
                    writer.WriteLine(s);
                }
            }
            Console.WriteLine("\nCSV file has been successfully created.");
            if (writer != null)
            {
                writer.Close();
                inputStream.Close();
                excelReader.Close();
            }
        }

    }
    class ConvertExec
    {
        static StreamWriter logWriter;
        static String logfilePath;
        static void Main(string[] args)
        {

            string directoryname = CheckFile();
            //string directoryname = "D:\\Temp";
            logfilePath = directoryname + "\\watch.log";

            var watcher = new FileSystemWatcher(directoryname);

            watcher.NotifyFilter = NotifyFilters.Attributes
                                 | NotifyFilters.CreationTime
                                 | NotifyFilters.DirectoryName
                                 | NotifyFilters.FileName
                                 | NotifyFilters.LastAccess
                                 | NotifyFilters.LastWrite
                                 | NotifyFilters.Security
                                 | NotifyFilters.Size;

            watcher.Changed += OnChanged;
            watcher.Created += OnCreated;
            watcher.Deleted += OnDeleted;
            watcher.Renamed += OnRenamed;
            watcher.Error += OnError;

            watcher.Filter = "*.csv";
            watcher.IncludeSubdirectories = true;
            watcher.EnableRaisingEvents = true;

            foreach (var files in Directory.GetFiles(directoryname))
            {
                FileInfo info = new FileInfo(files);
                var fileName = Path.GetFileName(info.FullName);
                Console.WriteLine(fileName);
                Console.WriteLine(info.Length);
                fileName = directoryname + "\\" + fileName;
                /*
                string conString = string.Empty;
                string storedProc = string.Empty;
                string sheet1 = string.Empty;
                string extension = Path.GetExtension(fileName);
                switch (extension)
                {
                    case ".xls": //Excel 97-03.
                        storedProc = "spx_ImportFromExcel03";
                        conString = ConfigurationManager.ConnectionStrings["Excel03ConString"].ConnectionString;
                        break;
                    case ".xlsx": //Excel 07 or higher.
                        storedProc = "spx_ImportFromExcel07";
                        conString = ConfigurationManager.ConnectionStrings["Excel07+ConString"].ConnectionString;
                        break;

                }
                
                //Read the Sheet Name.
                conString = string.Format(conString, fileName);
                using (OleDbConnection excel_con = new OleDbConnection(conString))
                {
                    excel_con.Open();
                    sheet1 = excel_con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null).Rows[0]["TABLE_NAME"].ToString();
                    excel_con.Close();
                }


                //Call the Stored Procedure to import Excel data in Table.
                string constr = ConfigurationManager.ConnectionStrings["conString"].ConnectionString;
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(storedProc, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@SheetName", sheet1);
                        cmd.Parameters.AddWithValue("@FilePath", fileName);
                        cmd.Parameters.AddWithValue("@HDR", "YES");
                        cmd.Parameters.AddWithValue("@TableName", "TBEXAMPLE");
                        cmd.Connection = con;
                        con.Open();
                        cmd.ExecuteNonQuery();
                        con.Close();
                    }
                }
                */
                if (!File.Exists(fileName) || (Path.GetExtension(fileName) != ".xls" && Path.GetExtension(fileName) != ".xlsx"))
                {
                    Console.WriteLine("\nInvalid file path or extension.");
                }
                else
                {

                    try
                    {
                        if (fileName != null)
                        {
                            ExcelToCsv obj = new ExcelToCsv(fileName);
                            string opfilename = fileName.Substring(0, (fileName.IndexOf(".xls"))) + ".csv";
                            obj.Convert(opfilename);
                            DataTable dataTable = GetDataTabletFromCSVFile(opfilename);
                            InsertDataIntoSQLServerUsingSQLBulkCopy(dataTable);

                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("\nAn exception has occured.");
                        Console.WriteLine(e.ToString());
                    }
                }
            }
            Console.WriteLine("Terminating...");
            Console.ReadLine();
        }
        private static string CheckFile()
        {
            Console.Write("\nEnter \\path: ");
            string fileName = Console.ReadLine();
            fileName = fileName.Replace(@"\", @"\\");
            fileName = fileName.Replace(@"/", @"\\");
            return fileName;
            /*
            // Check if file exists and file type is supported
            if (!File.Exists(fileName) || (Path.GetExtension(fileName) != ".xls" && Path.GetExtension(fileName) != ".xlsx"))
            {
                Console.WriteLine("\nInvalid file path or extension.");
                return null;
            }
            else
                return fileName;
            */
        }
        private static DataTable GetDataTabletFromCSVFile(string csv_file_path)
        {
            DataTable csvData = new DataTable();
            try
            {
                using (TextFieldParser csvReader = new TextFieldParser(csv_file_path))
                {
                    csvReader.SetDelimiters(new string[] { "," });
                    csvReader.HasFieldsEnclosedInQuotes = true;
                    string[] colFields = csvReader.ReadFields();
                    foreach (string column in colFields)
                    {
                        DataColumn datecolumn = new DataColumn(column);
                        datecolumn.AllowDBNull = true;
                        csvData.Columns.Add(datecolumn);
                    }
                    while (!csvReader.EndOfData)
                    {
                        string[] fieldData = csvReader.ReadFields();
                        //Making empty value as null
                        for (int i = 0; i < fieldData.Length; i++)
                        {
                            if (fieldData[i] == "")
                            {
                                fieldData[i] = null;
                            }
                        }
                        csvData.Rows.Add(fieldData);
                    }
                }
            }
            catch (Exception ex)
            {
                return null;
            }
            return csvData;
        }
        private static void InsertDataIntoSQLServerUsingSQLBulkCopy(DataTable csvFileData)
        {
            using (SqlConnection dbConnection = new SqlConnection("Data Source=localhost;database=testdb;Integrated Security=true"))
            {
                dbConnection.Open();
                using (SqlBulkCopy s = new SqlBulkCopy(dbConnection))
                {
                    s.DestinationTableName = "tblPersons";
                    foreach (var column in csvFileData.Columns)
                        s.ColumnMappings.Add(column.ToString(), column.ToString());
                    s.WriteToServer(csvFileData);
                }
            }
        }
        private static void OnChanged(object sender, FileSystemEventArgs e)
        {
            if (e.ChangeType != WatcherChangeTypes.Changed)
            {
                return;
            }
            logWriter = new StreamWriter(logfilePath, true);

            logWriter.WriteLine($"{DateTime.UtcNow.ToString()} Changed: {e.FullPath}");

            Console.WriteLine($"{DateTime.UtcNow.ToString()} Changed: {e.FullPath}");

            if (logWriter != null)
                logWriter.Close();
        }

        private static void OnCreated(object sender, FileSystemEventArgs e)
        {
            string value = $"{DateTime.UtcNow.ToString()} Created: {e.FullPath}";
            logWriter = new StreamWriter(logfilePath, true);
            logWriter.WriteLine(value);

            Console.WriteLine(value);

            if (logWriter != null)
                logWriter.Close();
        }

        private static void OnDeleted(object sender, FileSystemEventArgs e)
        {
            logWriter = new StreamWriter(logfilePath, true);
            logWriter.WriteLine($"{DateTime.UtcNow.ToString()} Deleted: {e.FullPath}");

            Console.WriteLine($"{DateTime.UtcNow.ToString()} Deleted: {e.FullPath}");

            if (logWriter != null)
                logWriter.Close();
        }
        private static void OnRenamed(object sender, RenamedEventArgs e)
        {
            logWriter = new StreamWriter(logfilePath, true);
            logWriter.WriteLine($"{DateTime.UtcNow.ToString()} Renamed:");
            logWriter.WriteLine($"    Old: {e.OldFullPath}");
            logWriter.WriteLine($"    New: {e.FullPath}");

            Console.WriteLine($"{DateTime.UtcNow.ToString()} Renamed:");
            Console.WriteLine($"    Old: {e.OldFullPath}");
            Console.WriteLine($"    New: {e.FullPath}");

            if (logWriter != null)
                logWriter.Close();
        }

        private static void OnError(object sender, ErrorEventArgs e) =>
            PrintException(e.GetException());

        private static void PrintException(Exception ex)
        {
            if (ex != null)
            {
                logWriter = new StreamWriter(logfilePath, true);
                logWriter.WriteLine($"{DateTime.UtcNow.ToString()} Message: {ex.Message}");
                logWriter.WriteLine("Stacktrace:");
                logWriter.WriteLine(ex.StackTrace);
                logWriter.WriteLine();

                if (logWriter != null)
                    logWriter.Close();
                Console.WriteLine($"{DateTime.UtcNow.ToString()} Message: {ex.Message}");
                Console.WriteLine("Stacktrace:");
                Console.WriteLine(ex.StackTrace);
                Console.WriteLine();
                PrintException(ex.InnerException);
            }
        }
        
    }
}

