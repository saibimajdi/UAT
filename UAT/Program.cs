using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace UAT
{
    class Program
    {
        // get Robot Codes file full path from APP SETTING file
        public static readonly string codeRobotFilePath = System.Configuration.ConfigurationManager.AppSettings["CODEROBOT_FILEPATH"].ToString();

        // get Robot Codes from .xlsx file
        public static List<CodeRobotItem> CodeRobotItems = GetRobotCodes(codeRobotFilePath);

        // get destination folder path from APP SETTING
        public static string DestinationFolderPath = System.Configuration.ConfigurationManager.AppSettings["PATHDEST"].ToString();

        static void Main(string[] args)
        {
            // get source folders path from APP SETTING file
            var sourceFoldersPath = new List<string>()
            {
                System.Configuration.ConfigurationManager.AppSettings["PATH_FOLDER1"].ToString()
                //System.Configuration.ConfigurationManager.AppSettings["PATH_FOLDER2"].ToString(),
                //System.Configuration.ConfigurationManager.AppSettings["PATH_FOLDER3"].ToString(),
                //System.Configuration.ConfigurationManager.AppSettings["PATH_FOLDER4"].ToString(),
            };

            FindMatching(sourceFoldersPath);
            
            Console.ReadKey();
        }

        private static void FindMatching(List<string> sourceFoldersPath)
        {
            if (sourceFoldersPath == null)
                return;

            var index = 1;

            foreach(var folderPath in sourceFoldersPath)
            {
                Console.WriteLine($"Checking folder [{index}; Path={folderPath}]...");
                var filesPaths = Directory.GetFiles(folderPath).ToList();
                FindMatchingForSpecificFolder(filesPaths);
                index++;
            }
        }

        private static void FindMatchingForSpecificFolder(List<string> filesPaths)
        {
            if (filesPaths == null)
                return;

            var index = 1;

            foreach (var filePath in filesPaths)
            {
                // console log
                Console.Write($"Checking file [{index}] ==> ");

                // check if this file contain any robot_code
                var codeRobot = FileContainsRobotCode(filePath);

                if (codeRobot != null)
                {
                    // console log
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine($"[YES] File path:{filePath}");
                    Console.ResetColor();

                    // generate new file name
                    var copiedFilename = $"{codeRobot.CodeRobot}_{codeRobot.NumSapClient}_{Path.GetFileName(filePath)}";

                    // copy the file from the source folder to the destination one
                    File.Copy(filePath, $"{DestinationFolderPath}/{copiedFilename}");
                }
                else
                {
                    // console log
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("[NO]");
                    Console.ResetColor();
                }
                index++;
            }
        }

        private static CodeRobotItem FileContainsRobotCode(string filePath)
        {
            foreach(var robotCode in CodeRobotItems)
            {
                var code = $"EDI+{robotCode.CodeRobot}:EDI";
                var fileContent = File.ReadAllText(filePath);
                if (fileContent.Contains(code))
                    return robotCode;
            }

            return null;
        }

        #region Excel Helpers
        static List<CodeRobotItem> GetRobotCodes(string filePath)
        {
            Console.WriteLine("Reading data from RobotCodes (.xlsx) file...");

            List<CodeRobotItem> codes = new List<CodeRobotItem>();

            Excel.Application xlApp = new Excel.Application();

            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);

            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];

            Excel.Range xlRange = xlWorksheet.UsedRange;

            for (int index = 1; index <= xlRange.Rows.Count; index++)
            {
                //Console.WriteLine($"[{index}] :{xlRange.Cells[index, 1].Value2.ToString()}");
                codes.Add(new CodeRobotItem(xlRange.Cells[index, 1].Value2.ToString(),
                                            xlRange.Cells[index, 2].Value2.ToString(),
                                            xlRange.Cells[index, 3].Value2.ToString()));
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorksheet);

            xlWorkbook.Close();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook);

            xlApp.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

            Console.WriteLine("Reading data from RobotCodes (.xlsx) file... [Completed]");

            return codes;
        }
        #endregion
    }

    public class CodeRobotItem
    {
        public string CodeRobot { get; set; }
        public string NumSapClient { get; set; }
        public string Canal { get; set; }
        public List<string> Log { get; set; } 
        
        public CodeRobotItem()
        {
            Log = new List<string>(); 
        }

        public CodeRobotItem(string codeRobot, string numSapClient, string canal)
        {
            CodeRobot = codeRobot;
            NumSapClient = numSapClient;
            Canal = canal;
            Log = new List<string>();
        }
    }
}
