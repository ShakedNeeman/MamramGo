using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ShowApp
{
    internal class Helper
    {
       

        public static Dictionary<string, Tuple<DateTime, string, int>> GetLog(int eventID)
        {
            // Initialize the dictionary to store the results
            Dictionary<string, Tuple< DateTime, string, int>> resultDict = new Dictionary<string, Tuple< DateTime, string, int>>();

            // Define the log we'll be reading from
            string logType = "Application"; // Adjust as needed

            // Read the event log
            EventLog eventLog = new EventLog(logType);

            // Go through each entry in the event log
            foreach (EventLogEntry logEntry in eventLog.Entries)
            {
                // Filter for Event ID  (Use InstanceId instead of deprecated EventID)
                if ((logEntry.EventID == eventID) && (eventID == 1034))
                {
                    // Add details to the dictionary
                    resultDict[GetNameFromMessageHebrew(logEntry.Message)] = Tuple.Create(logEntry.TimeGenerated, logEntry.Source, logEntry.EventID);
                }
                else if ((logEntry.EventID == eventID) && (eventID == 11707))
                {
                    resultDict[GetNameFromMessage(logEntry.Message)] = Tuple.Create(logEntry.TimeGenerated, logEntry.Source, logEntry.EventID);

                }
            }

            return resultDict;
        }

        public static string GetNameFromMessageHebrew(string message)
        {
            string productNamePrefix = "שם המוצר: ";
            string productNameSuffix = ". גירסת המוצר";

            int startIndex = message.IndexOf(productNamePrefix) + productNamePrefix.Length;
            int endIndex = message.IndexOf(productNameSuffix);

            if (startIndex > -1 && endIndex > -1)
            {
                string productName = message.Substring(startIndex, endIndex - startIndex);
                return productName;
            }
            else
            {
                return message;
            }

        }

        public static string GetNameFromMessage(string message)
        {
            // Hebrew prefixes and suffixes
            string productNamePrefixHeb = "מוצר: ";
            string productNameSuffixHeb = "-- ההתקנה הושלמה בהצלחה.";

            // English prefixes and suffixes
            string productNamePrefixEng = "Product: ";
            string productNameSuffixEng = " -- Installation completed successfully.";

            // Detect language and set prefix and suffix accordingly
            string productNamePrefix, productNameSuffix;

            if (message.Contains(productNamePrefixHeb))
            {
                productNamePrefix = productNamePrefixHeb;
                productNameSuffix = productNameSuffixHeb;
            }
            else
            {
                productNamePrefix = productNamePrefixEng;
                productNameSuffix = productNameSuffixEng;
            }

            int startIndex = message.IndexOf(productNamePrefix) + productNamePrefix.Length;
            int endIndex = message.IndexOf(productNameSuffix);

            if (startIndex > -1 && endIndex > -1)
            {
                string productName = message.Substring(startIndex, endIndex - startIndex);
                return productName;
            }
            else
            {
                return message;
            }
        }


        public static void CollectAppDetails(string path, Dictionary<string, Tuple<string, DateTime, long>> appDetails)
        {
            try
            {
                foreach (var dir in Directory.GetDirectories(path))
                {
                    var dirInfo = new DirectoryInfo(dir);

                   
                        string appName = dirInfo.Name;
                        DateTime installDate = dirInfo.CreationTime;
                        long size = GetDirectorySize(dirInfo);

                        if (appDetails.ContainsKey(appName))
                        {
                            if (appDetails[appName].Item3 != size)
                            {
                                appDetails[appName] = new Tuple<string, DateTime, long>(dir, installDate, size);
                            }
                        }
                        else
                        {
                            appDetails[appName] = new Tuple<string, DateTime, long>(dir, installDate, size);
                        }
                    
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"Could not access path: {path}, Error: {e.Message}");
            }
        }

        public static bool ContainsExecutable(DirectoryInfo dirInfo)
        {
            FileInfo[] files = dirInfo.GetFiles("*.exe", SearchOption.AllDirectories);
            return files.Length > 0;
        }

        public static long GetDirectorySize(DirectoryInfo dirInfo)
        {
            long size = 0;
            FileInfo[] fis = dirInfo.GetFiles();
            foreach (FileInfo fi in fis)
            {
                size += fi.Length;
            }

            DirectoryInfo[] dis = dirInfo.GetDirectories();
            foreach (DirectoryInfo di in dis)
            {
                size += GetDirectorySize(di);
            }

            return size;
        }
    }
}
