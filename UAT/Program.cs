using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UAT
{
    class Program
    {
        // get Robot Codes file full path from APP SETTING file
        public static readonly string codeRobotFilePath = System.Configuration.ConfigurationManager.AppSettings["CODEROBOT_FILEPATH"].ToString();

        // get Robot Codes from .xlsx file
        public static List<CodeRobotItem> CodeRobotItems = GetRobotCodesFromCSV(codeRobotFilePath);

        // get destination folder path from APP SETTING
        public static string DestinationFolderPath = System.Configuration.ConfigurationManager.AppSettings["PATHDEST"].ToString();

        static void Main(string[] args)
        {
            // get source folders path from APP SETTING file
            var sourceFoldersPath = new List<string>()
            {
                System.Configuration.ConfigurationManager.AppSettings["PATH_FOLDER1"].ToString(),
                System.Configuration.ConfigurationManager.AppSettings["PATH_FOLDER2"].ToString(),
                System.Configuration.ConfigurationManager.AppSettings["PATH_FOLDER3"].ToString(),
                System.Configuration.ConfigurationManager.AppSettings["PATH_FOLDER4"].ToString(),
            };

            // remove the first item that contains the columns name
            CodeRobotItems.RemoveAt(0);

            // summary container
            List<string> summary = new List<string>();

            // for each robot code, generate the specific files
            foreach (var codeItem in CodeRobotItems)
            {
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine($"Searching for files that contains the code: {codeItem.CodeRobot}...");
                Console.ResetColor();

                List<string> filesThatContainThisCode = new List<string>();

                // get the chosen files
                filesThatContainThisCode = FindMatching(sourceFoldersPath, codeItem);

                // console log
                Console.WriteLine($"{filesThatContainThisCode.Count} file(s) contain this code {codeItem.CodeRobot}");

                // sort the list of file by the creation time
                filesThatContainThisCode.Sort(
                    (path1, path2) => (File.GetCreationTime(path2) - File.GetCreationTime(path1)).Seconds
                    );

                // new file name
                var generatedFileName = $"{codeItem.CodeRobot}_{codeItem.NumSapClient}_.EDI";

                // copy the first file to the destination folder
                if (filesThatContainThisCode != null && filesThatContainThisCode.Count > 0)
                {
                    summary.Add($"{filesThatContainThisCode.Count} file(s) contains the {codeItem.CodeRobot}!");
                    
                    // delete the file if exist
                    if (File.Exists($"{DestinationFolderPath}/{generatedFileName}"))
                        File.Delete($"{DestinationFolderPath}/{generatedFileName}");

                    File.Copy(filesThatContainThisCode.FirstOrDefault(), $"{DestinationFolderPath}/{generatedFileName}");

                    // get edited BGMs
                    string oldBGM = string.Join(";", EditFileBGM($"{DestinationFolderPath}/{generatedFileName}"));

                    // log to file
                    codeItem.Log.Add($"[SOURCE]={filesThatContainThisCode.FirstOrDefault()};[DESTINATION]={DestinationFolderPath}/{generatedFileName};{oldBGM}");
                }
                else
                {
                    // log to file
                    codeItem.Log.Add($"[SOURCE]=NULL;[DESTINATION]=NULL");
                }
            }

            // check if the output file does exit
            if(!File.Exists($"{DestinationFolderPath}/output.csv"))
                File.Create($"{DestinationFolderPath}/output.csv").Close();

            // create the output file
            WriteRobotCodesToCSV(CodeRobotItems, $"{DestinationFolderPath}/output.csv");


            // display summary
            if(summary.Count > 0)
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("SUMMARY:");
                foreach (var str in summary)
                    Console.WriteLine(str);
                Console.ResetColor();
            }
            Console.WriteLine("Please press any key to continue..");
            Console.ReadKey();
        }

        private static List<string> EditFileBGM(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || string.IsNullOrEmpty(path))
                return null;

            string[] lines = File.ReadAllLines(path);

            List<string> oldBGM = new List<string>();

            for (int index = 0; index < lines.Length; ++index)
            {
                if((lines[index].StartsWith("BGM+220+"))&& (lines[index].IndexOf("9")>0))
                {
                    // saving old BGMs
                    oldBGM.Add(lines[index]);

                    lines[index] = lines[index].Replace("BGM+220+", "BGM+220+UAT_");
                }
            }

            // write lines on the file
            using (var writer = new StreamWriter(path, false)) // false to replace the file content (not append)
            {
                for (int index = 0; index < lines.Length; ++index)
                {
                    writer.WriteLine(lines[index]);
                }

                writer.Close();
            }

            return oldBGM;
        }

        // closing issue #1

        private static List<string> FindMatching(List<string> sourceFoldersPath, CodeRobotItem codeItem)
        {
            if (sourceFoldersPath == null)
                return null;

            var index = 1;

            List<string> filesThatContainThisCode = new List<string>();

            foreach (var folderPath in sourceFoldersPath)
            {
                // invalid folder path
                if (string.IsNullOrEmpty(folderPath) || string.IsNullOrWhiteSpace(folderPath))
                    return filesThatContainThisCode;

                Console.WriteLine($"Checking folder [{index}; Path={folderPath}]...");
                var filesPaths = Directory.GetFiles(folderPath).ToList();
                filesThatContainThisCode = FindMatchingForSpecificFolder(filesPaths, codeItem);

                // folder contains files that contains the robot code!
                if (filesThatContainThisCode.Count > 0)
                    return filesThatContainThisCode;

                index++;
            }

            return filesThatContainThisCode;
        }

        private static List<string> FindMatchingForSpecificFolder(List<string> filesPaths, CodeRobotItem codeItem)
        {
            if (filesPaths == null)
                return null;

            var index = 1;

            List<string> filesThatContainThisCode = new List<string>();

            foreach (var filePath in filesPaths)
            {
                // console log
                //Console.Write($"Checking file [{index}] ==> ");

                // check if this file contain any robot_code
                if (FileContainsRobotCode(filePath, codeItem))
                {
                    // console log
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine($"[YES] File path:{filePath}");
                    Console.ResetColor();

                    filesThatContainThisCode.Add(filePath);
                }
                //else
                //{
                //    // console log
                //    Console.ForegroundColor = ConsoleColor.Red;
                //    Console.WriteLine("[NO]");
                //    Console.ResetColor();
                //}
                index++;
            }

            return filesThatContainThisCode;
        }

        //private static bool FileContainsRobotCode(string filePath, CodeRobotItem codeItem)
        //{
        //    var code = $"EDI+{codeItem.CodeRobot}:EDI";
        //    var fileContent = File.ReadAllText(filePath);
        //    if (fileContent.Contains(code))
        //        return true;

        //    return false;
        //}

        // to be verified
        private static bool FileContainsRobotCode(string filePath, CodeRobotItem codeItem)
        {
            string fileContent = File.ReadAllText(filePath);

            /*var code = $"EDI+{codeItem.CodeRobot}:EDI";
            List<string> fileContent = File.ReadAllLines(filePath).ToList();
            foreach (string s in fileContent)
            {
                if (((s.IndexOf("UNB+UNOA:")) > 0) && (s.IndexOf(":EDI") > 0))
                {
                    string res = s.Substring(s.IndexOf("UNB+UNOA:"), s.IndexOf(":EDI"));
                    if (res.Contains(codeItem.CodeRobot))
                    {
                        return true;
                    }
                }
            }*/
            if ((fileContent.Contains("UNB+UNOA:")) && (fileContent.Contains(codeItem.CodeRobot)))
                return true;

            return false;
        }

        #region Excel Helpers

        static List<CodeRobotItem> GetRobotCodesFromCSV(string filePath)
        {
            List<CodeRobotItem> codes = new List<CodeRobotItem>();

            using (var fileStream = File.OpenRead(filePath))
            {
                using (var streamReader = new StreamReader(fileStream))
                {
                    while(!streamReader.EndOfStream)
                    {
                        var line = streamReader.ReadLine();
                        var values = line.Split(';');

                        codes.Add(new CodeRobotItem(values[0], values[1], values[2]));
                    }
                }
            }

            return codes;
        }
        static void WriteRobotCodesToCSV(List<CodeRobotItem> codes, string filePath)
        {

            using (var fileStream = File.OpenWrite(filePath))
            {
                using (var writer = new StreamWriter(fileStream))
                {
                    writer.WriteLine("code_robot; num_sap_client; canal; log");

                    foreach(var code in codes)
                    {
                        var log = string.Join(":::", code.Log.ToArray());

                        var line = $"{code.CodeRobot};{code.NumSapClient};{code.Canal};{log}";

                        writer.WriteLine(line);
                    }

                    writer.Close();
                }

                fileStream.Close();
            }
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
