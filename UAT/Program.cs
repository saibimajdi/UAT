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

        // start date
        public static readonly DateTime StartDate = DateTime.Parse(System.Configuration.ConfigurationManager.AppSettings["START_DATE"].ToString());

        // end date
        public static readonly DateTime EndDate = DateTime.Parse(System.Configuration.ConfigurationManager.AppSettings["END_DATE"].ToString());

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

            // temp code items list
            List<CodeRobotItem> _codeItemList = new List<CodeRobotItem>();

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
                //var generatedFileName = $"{codeItem.CodeRobot}_{codeItem.NumSapClient}_.EDI";



                // copy the first file to the destination folder
                if (filesThatContainThisCode != null && filesThatContainThisCode.Count > 0)
                {
                    // new file name
                    var filename = Path.GetFileName(filesThatContainThisCode.FirstOrDefault());

                    summary.Add($"{filesThatContainThisCode.Count} file(s) contains the {codeItem.CodeRobot}!");
                    
                    // delete the file if exist
                    if (File.Exists($"{DestinationFolderPath}\\{filename}"))
                        File.Delete($"{DestinationFolderPath}\\{filename}");

                    File.Copy(filesThatContainThisCode.FirstOrDefault(), $"{DestinationFolderPath}\\{filename}");

                    // get edited BGMs
                    var listOldBGM = EditFileBGM($"{DestinationFolderPath}\\{filename}");

                    if(listOldBGM != null && listOldBGM.Count > 0)
                    {
                        // log to file
                        foreach (var bgm in listOldBGM)
                        {
                            //codeItem.Log.Add($"{filesThatContainThisCode.FirstOrDefault()};{oldBGM}");
                            var _codeItem = new CodeRobotItem(codeRobot: codeItem.CodeRobot,
                                                           numSapClient: codeItem.NumSapClient, 
                                                                  canal: codeItem.Canal, 
                                                                 source: $"{filesThatContainThisCode.FirstOrDefault()}", 
                                                               commande: GetCommandeFromBGM(bgm), 
                                                           creationDate: File.GetCreationTime($"{filesThatContainThisCode.FirstOrDefault()}").ToShortDateString());

                            //_codeItem.Log.Add($"{filesThatContainThisCode.FirstOrDefault()};{bgm}");

                            _codeItemList.Add(_codeItem);
                        }
                    }
                    else
                    {
                        var _codeItem = new CodeRobotItem(codeRobot: codeItem.CodeRobot,
                                                       numSapClient: codeItem.NumSapClient,
                                                              canal: codeItem.Canal,
                                                             source: $"{filesThatContainThisCode.FirstOrDefault()}",
                                                           commande: "NULL",
                                                       creationDate: File.GetCreationTime($"{filesThatContainThisCode.FirstOrDefault()}").ToShortDateString());

                        _codeItemList.Add(_codeItem);
                    }
                }
                //else
                //{
                //    // log to file
                //    codeItem.Log.Add($"NULL;");
                //}
            }

            // check if the output file does exit
            if(!File.Exists($"{DestinationFolderPath}\\output.csv"))
                File.Create($"{DestinationFolderPath}\\output.csv").Close();

            // create the output file
            WriteRobotCodesToCSV(_codeItemList, $"{DestinationFolderPath}\\output.csv");


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

        private static string GetCommandeFromBGM(string bgm)
        {
            string commande = "";
            if(bgm.StartsWith("BGM+220+")) // BGM+220+
            {
                commande = bgm.Substring("BGM+220+".Length, bgm.Substring("BGM+220+".Length, (bgm.Length - "BGM+220+".Length)).IndexOf("+"));
            }
            else // BGM+220+
            {
                commande = bgm.Substring("BGM+105+".Length, bgm.Substring("BGM+105+".Length, (bgm.Length - "BGM+220+".Length)).IndexOf("+"));
            }

            return $"UAT_{commande}";
        }

        private static List<string> EditFileBGM(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || string.IsNullOrEmpty(path))
                return null;

            string[] lines = File.ReadAllLines(path);

            List<string> oldBGM = new List<string>();

            for (int index = 0; index < lines.Length; ++index)
            {
                if ((lines[index].StartsWith("BGM+220+")) && (lines[index].IndexOf("9") > 0))
                {
                    // saving old BGMs
                    oldBGM.Add(lines[index]);

                    lines[index] = lines[index].Replace("BGM+220+", "BGM+220+UAT_");
                }
                else
                {
                    if (lines[index].StartsWith("BGM+105+"))
                    {
                        oldBGM.Add(lines[index]);
                        lines[index] = lines[index].Replace("BGM+105+", "BGM+105+UAT_");
                    }
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
                // check creation date 
                /* var creationDate = File.GetLastWriteTime(filePath);
                 int result1 = DateTime.Compare(creationDate, StartDate);
                 int result2 = DateTime.Compare(creationDate, EndDate.AddHours(24));

                 if ((result1 < 1) || (result2>0))
                     break;*/

                // check if this file contain any robot_code
                if (FileContainsRobotCode(filePath, codeItem))
                {
                    // console log
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine($"[YES] File path:{filePath}");
                    Console.ResetColor();

                    filesThatContainThisCode.Add(filePath);
                }
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
            var code1 = $"UNB+UNOA:1+{codeItem.CodeRobot}:";
            var code2 = $"UNB+UNOA:3+{codeItem.CodeRobot}:";
            if ((fileContent.Contains(code1)) || (fileContent.Contains(code2)))
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

                        codes.Add(new CodeRobotItem(values[0], values[1], values[2], "", "", ""));
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
                    writer.WriteLine("code_robot; num_sap_client; canal; source; commande; date fichier");

                    foreach(var code in codes)
                    {
                        var line = $"{code.CodeRobot}; {code.NumSapClient}; {code.Canal}; {code.Source} ; {code.Commande}; {code.CreationDate}";

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
        public string Source { get; set; } 
        public string Commande { get; set; }
        public string CreationDate { get; set; }
        
        public CodeRobotItem(){}

        public CodeRobotItem(string codeRobot, string numSapClient, string canal, string source, string commande, string creationDate)
        {
            CodeRobot = codeRobot;
            NumSapClient = numSapClient;
            Canal = canal;
            Source = source;
            Commande = commande;
            CreationDate = creationDate;
        }
    }
}
