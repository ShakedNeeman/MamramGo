using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ShowApp
{
     class WinApp
    {
        public string AppName { get; set; }
        public string? Version { get; set; }
        public string? PackageCode { get; set; }
        public string? InstallLocation { get; set; }
        public string? InstallState { get; set; }
        public DateTime? InstallDate { get; set; }
        public int? InstalledSize { get; set; }
        public int? EventID { get; set; }
        public string? Source { get; set; }



        public static void ToString(WinApp app)
        {
            Console.WriteLine($"AppName: {app.AppName ?? "NULL"}");
            Console.WriteLine($"Version: {app.Version ?? "NULL"}");
            Console.WriteLine($"PackageCode: {app.PackageCode ?? "NULL"}");
            Console.WriteLine($"InstallLocation: {app.InstallLocation ?? "NULL"}");
            Console.WriteLine($"InstallState: {app.InstallState ?? "NULL"}");
            Console.WriteLine($"InstallDate: {(app.InstallDate.HasValue ? app.InstallDate.Value.ToString() : "NULL")}");
            Console.WriteLine($"InstalledSize: {(app.InstalledSize.HasValue ? app.InstalledSize.Value.ToString() : "NULL")}");
            Console.WriteLine($"EventID: {(app.EventID.HasValue ? app.EventID.Value.ToString() : "NULL")}");
            Console.WriteLine($"Source: {app.Source ?? "NULL"}");
        }

        public static void PrintWinApps(WinApp[] apps)
        {
            int count = 1;
            foreach (var app in apps)
            {
                Console.WriteLine($"Application #{count}:");
                ToString(app);
                Console.WriteLine();
                count++;
            }

        }

        public static Dictionary<string, WinApp> MergeWinApps(params WinApp[][] winAppArrays)
        {
            Dictionary<string, WinApp> unifiedWinApps = new Dictionary<string, WinApp>();

            foreach (var winAppArray in winAppArrays)
            {
                foreach (var winApp in winAppArray)
                {
                    if (unifiedWinApps.TryGetValue(winApp.AppName, out var existingWinApp))
                    {
                        // Merge logic here
                        if (winApp.Version != null) existingWinApp.Version = winApp.Version;
                        if (winApp.PackageCode != null) existingWinApp.PackageCode = winApp.PackageCode;
                        if (winApp.InstallLocation != null) existingWinApp.InstallLocation = winApp.InstallLocation;
                        if (winApp.InstallState != null) existingWinApp.InstallState = winApp.InstallState;
                        if (winApp.InstallDate != null) existingWinApp.InstallDate = winApp.InstallDate;
                        if (winApp.InstalledSize != null) existingWinApp.InstalledSize = winApp.InstalledSize;
                        if (winApp.EventID != null) existingWinApp.EventID = winApp.EventID;
                        if (winApp.Source != null) existingWinApp.Source = winApp.Source;
                    }
                    else
                    {
                        unifiedWinApps[winApp.AppName] = winApp;
                    }
                }
            }

            return unifiedWinApps;
        }

        public static (Dictionary<string, WinApp> UpdatedApps, Dictionary<string, WinApp> RemovedApps) CompareAndCleanAppDictionaries(
      Dictionary<string, WinApp> installedApps,
      Dictionary<string, WinApp> uninstalledApps)
        {
            Dictionary<string, WinApp> updatedInstalledApps = new Dictionary<string, WinApp>(installedApps);
            Dictionary<string, WinApp> removedApps = new Dictionary<string, WinApp>();

            foreach (var uninstalledEntry in uninstalledApps)
            {
                string appName = uninstalledEntry.Key;
                WinApp uninstalledApp = uninstalledEntry.Value;

                if (updatedInstalledApps.TryGetValue(appName, out WinApp installedApp))
                {
                    DateTime? uninstalledDate = uninstalledApp.InstallDate;
                    DateTime? installedDate = installedApp.InstallDate;

                    if (uninstalledDate.HasValue && installedDate.HasValue)
                    {
                        if (installedDate < uninstalledDate)
                        {
                            updatedInstalledApps.Remove(appName);
                            removedApps[appName] = installedApp; 
                        }
                    }
                }
            }

            return (updatedInstalledApps, removedApps);
        }



        public static WinApp[] CreateFromWmiDetails(Dictionary<string, Tuple<string, string, string, string>> wmiAppDetails)
        {
            if (wmiAppDetails == null)
                return new WinApp[0];

            List<WinApp> winApps = new List<WinApp>();

            foreach (var kv in wmiAppDetails)
            {
                WinApp winApp = new WinApp
                {
                    AppName = kv.Key,
                    Version = kv.Value.Item1,
                    InstallLocation = kv.Value.Item2,
                    InstallState = kv.Value.Item3,
                };

                if (DateTime.TryParse(kv.Value.Item4, out DateTime installDate))
                {
                    winApp.InstallDate = installDate;
                }
                else
                {
                    winApp.InstallDate = DateTime.MinValue;
                }

                winApps.Add(winApp);
            }

            return winApps.ToArray();
        }


        public static WinApp[] CreateFromRegistry(Dictionary<string, Tuple<string, string >> registryAppDetails)
        {
            if (registryAppDetails == null)
                return new WinApp[0];

            return registryAppDetails.Select(kv => new WinApp
            {
                AppName = kv.Key,
                Version = kv.Value.Item1,
                InstallLocation = kv.Value.Item2,
                
            }).ToArray();
        }
        public static WinApp[] CreateFromAPI(Dictionary<string, Tuple<string, string, string, string>> registryAppDetails)
        {
            if (registryAppDetails == null)
                return new WinApp[0];

            List<WinApp> winApps = new List<WinApp>();

            foreach (var kv in registryAppDetails)
            {
                WinApp winApp = new WinApp
                {
                    AppName = kv.Key,
                    PackageCode = kv.Value.Item1,
                    InstallLocation = kv.Value.Item2,
                    InstallDate = DateTime.TryParse(kv.Value.Item3, out DateTime installDate) ? installDate : DateTime.MinValue,
                    InstalledSize = int.TryParse(kv.Value.Item4, out int installedSize) ? installedSize : 0
                };

                winApps.Add(winApp);
            }

            return winApps.ToArray();
        }
        public static WinApp[] CreateFromLogs(Dictionary<string, Tuple<DateTime, string, int>> myLog)
        {
            if (myLog == null)
                return new WinApp[0];

            List<WinApp> winApps = new List<WinApp>();

            foreach (var kv in myLog)
            {
                WinApp winApp = new WinApp
                {
                    AppName = kv.Key,
                    InstallDate = kv.Value.Item1,
                    Source = kv.Value.Item2,
                    EventID = kv.Value.Item3,
                
                };

                winApps.Add(winApp);
            }

            return winApps.ToArray();
        }

        public static WinApp[] CreateFromFolder(Dictionary<string, Tuple<string, DateTime, long>> myapps)
        {
            if (myapps == null)
                return new WinApp[0];

            List<WinApp> winApps = new List<WinApp>();

            foreach (var kv in myapps)
            {
                WinApp winApp = new WinApp
                {
                    AppName = kv.Key,
                    InstallDate = DateTime.TryParse(kv.Value.Item1, out DateTime installDate) ? installDate : DateTime.MinValue,
                };

                if (kv.Value.Item3 >= int.MinValue && kv.Value.Item3 <= int.MaxValue)
                {
                    winApp.InstalledSize = (int)kv.Value.Item3;
                }
                else
                {
                    winApp.InstalledSize = int.MinValue; 
                }

                winApps.Add(winApp);
            }

            return winApps.ToArray();
        }
        public static WinApp[] CreateFromUninstallApp(Dictionary<string, Tuple<string>> myapps)
        {
            if (myapps == null)
                return new WinApp[0];

            List<WinApp> winApps = new List<WinApp>();

            foreach (var kv in myapps)
            {
                WinApp winApp = new WinApp
                {
                    AppName = kv.Key,
                    InstallState = kv.Value.Item1,
                };

                

                winApps.Add(winApp);
            }

            return winApps.ToArray();
        }

        public static WinApp[] CreateFromStore(Dictionary<string, Tuple<DateTime, long>> myapps)
        {
            if (myapps == null)
                return new WinApp[0];

            List<WinApp> winApps = new List<WinApp>();

            foreach (var kv in myapps)
            {
                WinApp winApp = new WinApp
                {
                    AppName = kv.Key,
                    InstallDate = kv.Value.Item1,
                };

                if (kv.Value.Item2 >= int.MinValue && kv.Value.Item2 <= int.MaxValue)
                {
                    winApp.InstalledSize = (int)kv.Value.Item2;
                }
                else
                {
                    winApp.InstalledSize = int.MinValue;
                }

                winApps.Add(winApp);
            }

            return winApps.ToArray();
        }


  
            public static void ExportToExcel(Dictionary<string, WinApp> updatedApps, Dictionary<string, WinApp> removedApps)
            {
                using (var package = new ExcelPackage())
                {
                    // Create worksheet for updated apps
                    ExportToWorksheet(package, "UpdatedApps", updatedApps);

                    // Create worksheet for removed apps
                    ExportToWorksheet(package, "RemovedApps", removedApps);

                    FileInfo fileInfo = new FileInfo("WinAppsAfterProcessing.xlsx");
                    package.SaveAs(fileInfo);
                    Console.WriteLine($"Excel file has been saved at: {fileInfo.FullName}");
                }
            }

            private static void ExportToWorksheet(ExcelPackage package, string sheetName, Dictionary<string, WinApp> apps)
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetName);

                worksheet.Cells[1, 1].Value = "AppName";
                worksheet.Cells[1, 2].Value = "Version";
                worksheet.Cells[1, 3].Value = "PackageCode";
                worksheet.Cells[1, 4].Value = "InstallLocation";
                worksheet.Cells[1, 5].Value = "InstallState";
                worksheet.Cells[1, 6].Value = "InstallDate";
                worksheet.Cells[1, 7].Value = "InstalledSize";
                worksheet.Cells[1, 8].Value = "EventID";
                worksheet.Cells[1, 9].Value = "Source";

            int row = 2;

                foreach (var entry in apps)
                {
                    var app = entry.Value;
                    worksheet.Cells[row, 1].Value = app.AppName;
                    worksheet.Cells[row, 2].Value = app.Version;
                    worksheet.Cells[row, 3].Value = app.PackageCode;
                    worksheet.Cells[row, 4].Value = app.InstallLocation;
                    worksheet.Cells[row, 5].Value = app.InstallState;
                    worksheet.Cells[row, 6].Value = app.InstallDate?.ToShortDateString();
                    worksheet.Cells[row, 7].Value = app.InstalledSize;
                    worksheet.Cells[row, 8].Value = app.EventID;
                    worksheet.Cells[row, 9].Value = app.Source;

                    row++;

                }
            }

        }
    }







