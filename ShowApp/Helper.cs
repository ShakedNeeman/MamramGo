using System.Diagnostics;
using System;
using System.Collections.Generic;
using Windows.Management.Deployment;

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
            HashSet<string> ignoreList = new HashSet<string>
        {
          "Common Files",
        "Microsoft Shared",
        "Program Files",
        "Program Files (x86)",
        "Windows",
        "System",
        "System32",
        "Users",
        "Temp",
        "AppData",
        "NetHood",
        "PrintHood",
        "Recent",
        "SendTo",
        "Start Menu",
        "Boot",
        };

            try
            {
                foreach (var dir in Directory.GetDirectories(path))
                {
                    var dirInfo = new DirectoryInfo(dir);
                    string appName = dirInfo.Name;

                    // Skip directories in the ignore list
                    if (ignoreList.Contains(appName)) continue;

                    DateTime installDate = dirInfo.CreationTime;
                    long size = GetDirectorySize(dirInfo);

                    if (ContainsExecutable(dirInfo))
                    {
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




        public static Dictionary<string, Tuple<DateTime, long>> CollectAppDetailsStore()
        {
            // Initialize a dictionary to store details of installed UWP applications.
            Dictionary<string, Tuple<DateTime, long>> appDetails = new Dictionary<string, Tuple<DateTime, long>>();

            try
            {
                // Instantiate the PackageManager class to access UWP package information.
                PackageManager packageManager = new PackageManager();

                // Retrieve all packages for the current user, focusing on the main app packages.
                // This excludes things like framework and resource packages.
                var packages = packageManager.FindPackagesForUserWithPackageTypes(string.Empty, PackageTypes.Main);

                foreach (var package in packages)
                {
                    // If the package is a framework, resource package, or bundle, skip it.
                    if (package.IsFramework || package.IsResourcePackage || package.IsBundle)
                    {
                        continue;
                    }

                    // Check if the installation directory of the package contains "WindowsApps".
                    if (package.InstalledLocation.Path.Contains("WindowsApps"))
                    {
                        // Get the application name.
                        string appName = package.Id.Name;

                        // Get the installation date.
                        DateTime installDate = package.InstalledDate.LocalDateTime;

                        // Placeholder for size as it's not easily accessible for UWP apps.
                        long size = 0;

                        // Update the dictionary with the new details.
                        if (appDetails.ContainsKey(appName))
                        {
                            if (appDetails[appName].Item2 != size)
                            {
                                appDetails[appName] = new Tuple<DateTime, long>(installDate, size);
                            }
                        }
                        else
                        {
                            appDetails[appName] = new Tuple<DateTime, long>(installDate, size);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                // Log any exceptions.
                Console.WriteLine($"Could not access the PackageManager: {e.Message}");
            }

            return appDetails;
        }


    }
}
