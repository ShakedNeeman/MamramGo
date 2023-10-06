using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.Runtime.InteropServices;
using PlasmaNGService.HealthCore;
using System.Management;
using System.Threading;
using System.IO;
using OfficeOpenXml;
using Microsoft.Win32;
using Windows.Management.Deployment;



namespace PlasmaNGService.HealthCore
{


    class WindowsApps
    {

        #region fields
        private string _name;
        private string _path;
        private string _publisher;
        private string _date;
        private string _version;
        #endregion



        public class AppsNameComparer : IEqualityComparer<WindowsApps>
        {
            public bool Equals(WindowsApps x, WindowsApps y)
            {
                if (x == null || y == null)
                {
                    return false;
                }

                if ((!string.IsNullOrEmpty(x._name)) && (!string.IsNullOrEmpty(y._name)))
                {
                    if (x._name == y._name)
                    {
                        if (string.IsNullOrEmpty(x._path) && string.IsNullOrEmpty(y._path))
                        {
                            return true; 
                        }

                        if (!string.IsNullOrEmpty(x._path) && !string.IsNullOrEmpty(y._path) && x._path == y._path)
                        {
                            return true;  
                        }
                    }
                }

                return false; 
            }

            public int GetHashCode(WindowsApps obj)
          {
                if (!string.IsNullOrEmpty(obj.Name))
                {
                    if (!string.IsNullOrEmpty(obj.Path))
                    {
                        return obj.Name.GetHashCode() ^ obj.Path.GetHashCode();
                    }
                    else
                    {
                        return obj.Name.GetHashCode();
                    }
                }

                return 0; 
            }





        }


        /// <summary>
        /// Checks which applications are installed on the computer using Registry, MSI and WMI methods.
        /// Checks for overlaps between the applications discovered by the methods.
        /// </summary>
        /// 
        //getter of properties
        public string Name { get => _name; set { _name = value; } }
        public string Path { get => _path; set { _path = value; } }
        public string Publisher { get => _publisher; set { _publisher = value; } }
        public string Date { get => _date; set { _date = value; } }
        public string Version { get => _version; set { _version = value; } }

        public WindowsApps()
        {

        }


        public WindowsApps(string name, string path, string publisher, string appdate, string version)
        {
            _name = name;
            _path = path;
            _publisher = publisher;
            _date = appdate;
            _version = version;
        }

        public List<WindowsApps> GetInstalledAppsRegistry(string path)
        {
            RegistryKey regKey = Registry.LocalMachine.OpenSubKey(path);
            string[] subKeyNames = regKey.GetSubKeyNames();
            List<WindowsApps> installedApps = new List<WindowsApps>();

            foreach (string subKeyName in subKeyNames)
            {
                RegistryKey subKey = regKey.OpenSubKey(subKeyName);
                string displayName = subKey.GetValue("DisplayName") as string;
                if (!string.IsNullOrEmpty(displayName))
                {

                    string _path = subKey.GetValue("InstallLocationInstallLocation") as string;
                    if (string.IsNullOrEmpty(_path))
                    {
                        _path = subKey.GetValue("InstallSource") as string;
                    }
                    string _publisher = subKey.GetValue("Publisher") as string;
                    string installDateString = subKey.GetValue("Installdate") as string;
                    string _date = null;
                    if (installDateString != null)
                    {
                        //DateTime installDate = DateTime.ParseExact(installDateString, "dd/MM/yyy hh:mm:ss", null, System.Globalization.DateTimeStyles.AdjustToUniversal);
                        //app._date = installDate.ToString();
                        try
                        {
                            DateTime installDate = DateTime.ParseExact(installDateString, "yyyyMMdd", null);
                            _date = installDate.ToString();
                        }
                        catch (FormatException)
                        {
                            Console.WriteLine($"Failed to parse date: {installDateString}");
                        }
                    }
                    string _version = subKey.GetValue("DisplayVersion") as string;
                    WindowsApps app = new WindowsApps(displayName, _path, _publisher, _date, _version);
                    installedApps.Add(app);
                }
            }
            return installedApps;
        }


        public List<WindowsApps> GetInstalledAppsWMI()
        {
            List<WindowsApps> installedApps = new List<WindowsApps>();
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Product");
            ManagementObjectCollection products = searcher.Get();

            foreach (ManagementObject product in products)
            {
                string displayName = product["Name"] as string;
                if (!string.IsNullOrEmpty(displayName))
                {

                    string _publisher = product["Vendor"] as string;
                    string _path = product["InstallLocation"] as string;
                    if (string.IsNullOrEmpty(_path))
                    {
                        _path = product["InstallSource"] as string;
                    }
                    string installDateStr = product["InstallDate"] as string;
                    string _date = null;
                    if (installDateStr != null)
                    {
                        //DateTime installDate = DateTime.ParseExact(installDateStr, "yyyyMMdd", null);
                        //app._date = installDate.ToString();
                        try
                        {
                            DateTime installDate = DateTime.ParseExact(installDateStr, "yyyyMMdd", null);
                            _date = installDate.ToString();
                        }
                        catch (FormatException)
                        {
                            Console.WriteLine($"Failed to parse date: {installDateStr}");
                        }
                    }
                    string _version = product["Version"] as string;
                    WindowsApps app = new WindowsApps(displayName, _path, _publisher, _date, _version);
                    installedApps.Add(app);
                }
            }
            return installedApps;
        }

        public static List<WindowsApps> CollectAppDetailsStore()
        {
            List<WindowsApps> appDetails = new List<WindowsApps>();

            try
            {
                PackageManager packageManager = new PackageManager();
                var packages = packageManager.FindPackagesForUserWithPackageTypes(string.Empty, PackageTypes.Main);

                foreach (var package in packages)
                {
                    if (package.IsFramework || package.IsResourcePackage || package.IsBundle)
                    {
                        continue;
                    }

                    if (package.InstalledLocation.Path.Contains("WindowsApps"))
                    {
                        string appName = package.DisplayName;
                        string publisher = package.PublisherDisplayName;
                        string installDateStr = package.InstalledDate.LocalDateTime.ToString();

                        WindowsApps app = new WindowsApps(appName, package.InstalledLocation.Path, publisher, installDateStr, package.Id.Version.ToString());
                        appDetails.Add(app);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"Could not access the PackageManager: {e.Message}");
            }

            return appDetails;
        }

        public static List<WindowsApps> GetLog(int eventID)
        {
            // Initialize the list to store the results
            List<WindowsApps> resultApps = new List<WindowsApps>();

            // Define the log we'll be reading from
            string logType = "Application"; // Adjust as needed

            // Read the event log
            EventLog eventLog = new EventLog(logType);

            // Go through each entry in the event log
            foreach (EventLogEntry logEntry in eventLog.Entries)
            {
                // Filter for Event ID (Use InstanceId instead of deprecated EventID)
                if ((logEntry.EventID == eventID) && (eventID == 1034))
                {
                    string name = GetNameFromMessageHebrew(logEntry.Message);
                    DateTime timeGenerated = logEntry.TimeGenerated;
                    string source = logEntry.Source;
                    int eventId = logEntry.EventID;
                    string version = (logEntry.Message); // הוספתי את הפונקציה הזו

                    // Create a WindowsApps instance and add it to the list
                    WindowsApps app = new WindowsApps(name, null, source, timeGenerated.ToString(), eventId.ToString());
                    resultApps.Add(app);
                }
                else if ((logEntry.EventID == eventID) && (eventID == 11707))
                {
                    string name = GetNameFromMessage(logEntry.Message);
                    DateTime timeGenerated = logEntry.TimeGenerated;
                    string source = logEntry.Source;
                    int eventId = logEntry.EventID;
                    string version = (logEntry.Message); // הוספתי את הפונקציה הזו


                    // Create a WindowsApps instance and add it to the list
                    WindowsApps app = new WindowsApps(name, null, source, timeGenerated.ToString(), eventId.ToString());
                    resultApps.Add(app);
                }
            }

            return resultApps;
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
            string productNameSuffixEngB = " -- Installation operation completed successfully.";

            // Detect language and set prefix and suffix accordingly
            string productNamePrefix, productNameSuffix;

            if (message.Contains(productNamePrefixHeb))
            {
                productNamePrefix = productNamePrefixHeb;
                productNameSuffix = productNameSuffixHeb;
            }
            else
            {
                if (message.Contains(productNameSuffixEng))
                {
                    productNamePrefix = productNamePrefixEng;
                    productNameSuffix = productNameSuffixEng;
                }
                else
                {
                    productNamePrefix = productNamePrefixEng;
                    productNameSuffix = productNameSuffixEngB;
                }
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


        public static List<WindowsApps> UninstallApp()
        {
            // Initialize a list to hold uninstalled application details
            List<WindowsApps> uninstalledApps = new List<WindowsApps>();

            // Open the Registry key where installed apps are listed
            RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall");

            // Loop through all subkeys to find uninstalled apps
            foreach (string subKeyName in key.GetSubKeyNames())
            {
                // Open each subkey to read its values
                RegistryKey subKey = key.OpenSubKey(subKeyName);

                // Check for null and proceed
                if (subKey != null)
                {
                    // Read app details from Registry values
                    object displayName = subKey.GetValue("DisplayName");
                    object installLocation = subKey.GetValue("InstallLocation");
                    object publisher = subKey.GetValue("Publisher");
                    object installDateStr = subKey.GetValue("InstallDate");
                    object displayVersion = subKey.GetValue("DisplayVersion");

                    string name = displayName as string;
                    string path = installLocation as string;
                    string _publisher = publisher as string;
                    string _date = null;
                    if (installDateStr != null)
                    {
                        try
                        {
                            DateTime installDate = DateTime.ParseExact(installDateStr as string, "yyyyMMdd", null);
                            _date = installDate.ToString();
                        }
                        catch (FormatException)
                        {
                            Console.WriteLine($"Failed to parse date: {installDateStr}");
                        }
                    }
                    string version = displayVersion as string;

                    // Create a WindowsApps instance and add it to the list
                    WindowsApps app = new WindowsApps(name, path, _publisher, _date, version);
                    uninstalledApps.Add(app);
                }
            }

            // Return the list containing uninstalled app details
            return uninstalledApps;
        }

       
        public static List<WindowsApps> CollectAppDetails()
        {
                    List<string> paths = new List<string>
            {
                @"C:\Program Files",
                @"C:\Program Files (x86)",
                @"C:\Windows\System32",
                @"C:\ProgramData",
                @"C:\Users",
            };

                    List<WindowsApps> installedApps = new List<WindowsApps>();
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
                    HashSet<string> targetDirectories = new HashSet<string>
            {
                "data",
                "legal",
                "lib",
                "bin",
                "log",
                "setup"
            };

                    foreach (string path in paths)
                    {
                        int foundDirectoriesCount = 0;
                        int foundDLLsCount = 0;
                        int foundEXEsCount = 0;
                        int foundJARsCount = 0;

                        try
                        {
                            foreach (string itemPath in Directory.GetFileSystemEntries(path))
                            {
                                if (Directory.Exists(itemPath))
                                {
                                    // Check directories
                                    string dirName = System.IO.Path.GetFileName(itemPath);
                                    if (targetDirectories.Contains(dirName) && !ignoreList.Contains(dirName))
                                    {
                                        foundDirectoriesCount++;
                                        if (foundDirectoriesCount >= 2)
                                        {
                                    // We found at least 2 target directories, no need to continue checking
                                    var dirInfo = new DirectoryInfo(itemPath);
                                    string _date = dirInfo.CreationTime.ToString();
                                    string pathName = System.IO.Path.GetFullPath(itemPath);


                                    WindowsApps app = new WindowsApps(dirName, pathName, "", _date, "");
                                    installedApps.Add(app);
                                    break;
                                        }
                                    }
                                }
                                //else
                                //{
                                //    // Check files
                                //    string fileExtension = System.IO.Path.GetExtension(itemPath).ToLower();
                                //    if (fileExtension == ".dll")
                                //    {
                                //        foundDLLsCount++;
                                //    }
                                //    else if (fileExtension == ".exe")
                                //    {
                                //        foundEXEsCount++;
                                //    }
                                //    else if (fileExtension == ".jar")
                                //    {
                                //        foundJARsCount++;
                                //    }

                                //    if (foundDLLsCount >= 1 && foundEXEsCount >= 1 && foundJARsCount >= 1)
                                //    {
                                //        var dirInfo = new DirectoryInfo(itemPath);
                                //        string _date = dirInfo.CreationTime.ToString();
                                //        string pathName = System.IO.Path.GetFullPath(itemPath);


                                //        WindowsApps app = new WindowsApps(dirName, pathName, "", _date, "");
                                //installedApps.Add(app); break;
                                //    }
                                //}
                            }
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine($"Could not access path: {path}, Error: {e.Message}");
                        }
                    }

            return installedApps;
        }


        public static (List<WindowsApps> UpdatedApps, List<WindowsApps> RemovedApps) CompareAndCleanAppLists(
     List<WindowsApps> installedAppsList,
     List<WindowsApps> uninstalledAppsList)
        {
            List<WindowsApps> updatedAppsList = new List<WindowsApps>(installedAppsList);
            List<WindowsApps> removedAppsList = new List<WindowsApps>();

            foreach (var uninstalledApp in uninstalledAppsList)
            {
                WindowsApps installedApp = installedAppsList.FirstOrDefault(app => app._name == uninstalledApp._name);

                if (installedApp != null)
                {
                    string uninstalledDateStr = uninstalledApp._date;
                    string installedDateStr = installedApp._date;

                    if (!string.IsNullOrEmpty(uninstalledDateStr) && !string.IsNullOrEmpty(installedDateStr))
                    {
                        DateTime uninstalledDate = DateTime.Parse(uninstalledDateStr);
                        DateTime installedDate = DateTime.Parse(installedDateStr);


                        if (installedDate < uninstalledDate)
                        {
                            // The app was updated or reinstalled
                            updatedAppsList.Remove(installedApp);
                            removedAppsList.Add(uninstalledApp);
                        }
                    }
                }
            }

            return (updatedAppsList, removedAppsList);
        }




        private List<WindowsApps> GetCombinedInstalledList()
        {
            List<WindowsApps> registryApps = GetInstalledAppsRegistry("Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall");
            List<WindowsApps> registryApps64 = GetInstalledAppsRegistry("Software\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall");
            List<WindowsApps> WMIApps = GetInstalledAppsWMI();
            List<WindowsApps> WINStoreApps = CollectAppDetailsStore();
            List<WindowsApps> FolderApps = CollectAppDetails();
            List<WindowsApps> LogInstallApps = GetLog(11707);
            LogInstallApps = LogInstallApps.Distinct(new AppsNameComparer()).ToList();




            List<WindowsApps> combinedList = registryApps.Concat(registryApps64).ToList();
            combinedList = combinedList.Concat(WMIApps).ToList();
            combinedList = combinedList.Concat(WINStoreApps).ToList();
            combinedList = combinedList.Concat(FolderApps).ToList();
            combinedList = combinedList.Concat(LogInstallApps).ToList();

            return combinedList;
        }

        public static List<WindowsApps> GetUniqueApps(List<WindowsApps> apps)
        {
            List<WindowsApps> uniqueApps = new List<WindowsApps>();

            foreach (var app in apps)
            {
                bool isUnique = true;

                // Check if an app with the same name and different PATH exists
                foreach (var existingApp in uniqueApps)
                {
                    if (existingApp.Name == app.Name)
                    {
                        if (string.IsNullOrEmpty(existingApp.Path) || string.IsNullOrEmpty(app.Path))
                        {
                            isUnique = false;
                            break;
                        }
                        else if (!string.IsNullOrEmpty(existingApp.Path) && !string.IsNullOrEmpty(app.Path) && existingApp.Path != app.Path)
                        {
                            isUnique = false;
                            break;
                        }
                    }
                }

                if (isUnique)
                {
                    uniqueApps.Add(app);
                }
            }

            return uniqueApps;
        }

        private List<WindowsApps> GetCombinedUnInstalledList()
        {
     
            List<WindowsApps> LogUnInstallApps = GetLog(1034);
            List<WindowsApps> UnInstallApps = UninstallApp();


            List<WindowsApps> combinedList = LogUnInstallApps.Concat(UnInstallApps).ToList();
          

            return combinedList;
        }
        public List<WindowsApps> GetOverlapsApps()//need to change that will fit to the new Apps class
        {
            List<WindowsApps> installedApps = GetCombinedInstalledList();
            List<WindowsApps> distinct = installedApps.Distinct(new AppsNameComparer()).ToList();
            List<WindowsApps> overlappedApps = installedApps.Except(distinct).ToList();
            return overlappedApps;
        }


        public List<WindowsApps> GetInstalledApps()//need to change that will fit to the new Apps class
        {
            List<WindowsApps> installedApps = GetCombinedInstalledList();
            List<WindowsApps> distinct = installedApps.Distinct(new AppsNameComparer()).ToList();
            return distinct;
        }

        public List<WindowsApps> GetUnInstalledApps()//need to change that will fit to the new Apps class
        {
            List<WindowsApps> installedApps = GetCombinedUnInstalledList();
           // List<WindowsApps> distinct = installedApps.Distinct(new AppsNameComparer()).ToList();
            return installedApps;
        }

        public static void ExportToWorksheet(ExcelPackage package, string sheetName, List<WindowsApps> apps)
        {
            var worksheet = package.Workbook.Worksheets.Add(sheetName);

            worksheet.Cells[1, 1].Value = "Name";
            worksheet.Cells[1, 2].Value = "Path";
            worksheet.Cells[1, 3].Value = "Publisher";
            worksheet.Cells[1, 4].Value = "Date";
            worksheet.Cells[1, 5].Value = "Version";

            int row = 2;
            foreach (var app in apps)
            {
                worksheet.Cells[row, 1].Value = app.Name;
                worksheet.Cells[row, 2].Value = app.Path;
                worksheet.Cells[row, 3].Value = app.Publisher;
                worksheet.Cells[row, 4].Value = app.Date;
                worksheet.Cells[row, 5].Value = app.Version;
                row++;
            }
        }
    }

}