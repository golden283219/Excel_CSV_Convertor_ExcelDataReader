using System;
using System.IO;
using System.Data;
using Excel;
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
                writer.Close();
        }
    }
    class ConvertExec
    {
        static StreamWriter logWriter;
        static String logfilePath;
        static void Main(string[] args)
        {

            string filename = CheckFile();

            logfilePath = Path.GetDirectoryName(filename) + "\\watch.log";
            var watcher = new FileSystemWatcher(Path.GetDirectoryName(filename));

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

            try
            {
                if (filename != null)
                {
                    ExcelToCsv obj = new ExcelToCsv(filename);
                    string opfilename = filename.Substring(0, (filename.IndexOf(".xls"))) + ".csv";
                    obj.Convert(opfilename);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("\nAn exception has occured.");
                Console.WriteLine(e.ToString());
            }
            Console.WriteLine("Terminating...");
            Console.ReadLine();
        }
        private static string CheckFile()
        {
            Console.Write("\nEnter \\path\\to\\filename: ");
            string fileName = Console.ReadLine();
            fileName = fileName.Replace(@"\",@"\\");
            fileName = fileName.Replace(@"/", @"\\");
           // Check if file exists and file type is supported
            if (!File.Exists(fileName) || (Path.GetExtension(fileName) != ".xls" && Path.GetExtension(fileName) != ".xlsx"))
            {
                Console.WriteLine("\nInvalid file path or extension.");
                return null;
            }
            else
                return fileName;
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
