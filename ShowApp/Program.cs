
using Microsoft.Win32;
using System.Management;
using OfficeOpenXml;
using System.Diagnostics;
using ShowApp;
using System.Collections.Generic;
using PlasmaNGService.HealthCore;
using System.Collections;
using System.Xml;

class Program
{
    static void Main(string[] args)
    {
        // Your code here


bool continueMenu = true;
    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


    do
    {
        Console.WriteLine("Enter a number to select a display operation method");
        Console.WriteLine("1-My Plasma");
        Console.WriteLine("2-Plasma");
        Console.WriteLine("3-Export To CSV After After Processing");
        Console.WriteLine("4-Export To CSV ");



        Console.WriteLine("5-EXIT");

        int Number = int.Parse(Console.ReadLine());


        switch (Number)
        {
            case 1:
                    WindowsApps WindowsApps = new WindowsApps();


                    List<WindowsApps> distinctInstall = WindowsApps.GetInstalledApps();
                    List<WindowsApps> distinctUninstall = WindowsApps.GetUnInstalledApps();


                    var resultPlasma = WindowsApps.CompareAndCleanAppLists(distinctInstall, distinctUninstall);
                    var updatedAppsPlasma = resultPlasma.UpdatedApps;
                    var removedAppsPlasma = resultPlasma.RemovedApps;
                    updatedAppsPlasma = updatedAppsPlasma.Distinct(new WindowsApps.AppsNameComparer()).ToList();
                    removedAppsPlasma = removedAppsPlasma.Distinct(new WindowsApps.AppsNameComparer()).ToList();
                    var Unique = WindowsApps.GetUniqueApps(updatedAppsPlasma);


                    using (var package = new ExcelPackage())
                    {
                        WindowsApps.ExportToWorksheet(package, "Distinct", updatedAppsPlasma);
                        WindowsApps.ExportToWorksheet(package, "Removed", removedAppsPlasma);
                        WindowsApps.ExportToWorksheet(package, "Unique", Unique);


                        FileInfo fileInfo = new FileInfo("AppsDataByPlasmaD.xlsx");
                        package.SaveAs(fileInfo);
                        Console.WriteLine($"Excel file has been saved at: {fileInfo.FullName}");
                        Console.WriteLine($"Total distinctInstall apps found: {distinctInstall.Count}");
                        Console.WriteLine($"Total updatedAppsPlasma apps found: {updatedAppsPlasma.Count}");
                        Console.WriteLine($"Total Unique apps found: {Unique.Count}");




                    }


                    break;
            case 2:
                    WindowsApps windowsApps = new WindowsApps();

                    List<WindowsApps> registryApps = windowsApps.GetInstalledAppsRegistry("Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall");
                    Console.WriteLine($"Found {registryApps.Count} apps using Registry method.");

                    List<WindowsApps> registryApps64 = windowsApps.GetInstalledAppsRegistry("Software\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall");
                    Console.WriteLine($"Found {registryApps64.Count} apps using Registry 64 method.");

                    List<WindowsApps> WMIApps = windowsApps.GetInstalledAppsWMI();
                    Console.WriteLine($"Found {WMIApps.Count} apps using WMI method.");


                    List<WindowsApps> distinctApps = windowsApps.GetInstalledApps();
                    Console.WriteLine($"Total distinct apps found: {distinctApps.Count}");

                    List<WindowsApps> overlappedApps = windowsApps.GetOverlapsApps();
                    Console.WriteLine($"Total overlapped apps found: {overlappedApps.Count}");

                    break;
            case 3:
                    // Get data from each function
                    Dictionary<string, Tuple<string, string, string, string>> wmiAppDetailsAPP = WMI.GetAppByWMI();
                    Dictionary<string, Tuple<string,string>> registryAppDetailsAPP = MyRegistry.GetAppByRegistry("Software\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall");
                    Dictionary<string, Tuple<string, string>> registr = MyRegistry.GetAppByRegistry("Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall");
                    foreach (var pair in registr)
                    {
                        if (!registryAppDetailsAPP.ContainsKey(pair.Key))
                        {
                            registryAppDetailsAPP[pair.Key] = pair.Value;
                        }
                    }
                    Dictionary<string, Tuple<string, string, string, string>> appsByAppNameAPP = WinApi.GetAppsUsingAPI();
                    Dictionary<string,Tuple<string>> uninstalledApp = MyRegistry.UninstallApp();
                    Dictionary<string, Tuple<DateTime, string, int>> myLog1034APP = Helper.GetLog(1034);
                    Dictionary<string, Tuple<DateTime, string, int>> myLog11707APP = Helper.GetLog(11707);
                    Dictionary<string, Tuple<DateTime, long>> appFromStoreAPP = Helper.CollectAppDetailsStore();


                    Dictionary<string, Tuple<string, DateTime, long>> appDetailsAPP = new Dictionary<string, Tuple<string, DateTime, long>>();
                    Helper.CollectAppDetails("C:\\Program Files", appDetailsAPP);
                    Helper.CollectAppDetails("C:\\Program Files (x86)", appDetailsAPP);
                    Helper.CollectAppDetails($"C:\\Users\\{Environment.UserName}\\AppData\\", appDetailsAPP);
                    Helper.CollectAppDetails("C:\\Windows\\System32", appDetailsAPP);
                    Helper.CollectAppDetails("C:\\ProgramData", appDetailsAPP);

                    WinApp[] wmiApps = WinApp.CreateFromWmiDetails(wmiAppDetailsAPP);
                    WinApp[] registryApp = WinApp.CreateFromRegistry(registryAppDetailsAPP);
                    WinApp[] apiApps = WinApp.CreateFromAPI(appsByAppNameAPP);
                    WinApp[] logApps1034 = WinApp.CreateFromLogs(myLog1034APP);
                    WinApp[] logApps11707 = WinApp.CreateFromLogs(myLog11707APP);
                    WinApp[] AppStore = WinApp.CreateFromFolder(appDetailsAPP);
                    WinApp[] uninstalled = WinApp.CreateFromUninstallApp(uninstalledApp);
                    WinApp[] appFromStores = WinApp.CreateFromStore(appFromStoreAPP);

                    Dictionary<string, WinApp> mergedUninstalledApps = WinApp.MergeWinApps(
                       
                        logApps1034,
                         uninstalled
                       
                    );

                    Console.WriteLine("total app that installed" + mergedUninstalledApps.Count);

                    Dictionary<string, WinApp> mergedInstalledApps = WinApp.MergeWinApps(
                        registryApp,
                        wmiApps,
                        apiApps,
                        logApps11707,
                        AppStore,
                        appFromStores
                    );
                    Console.WriteLine("total app just installed appDetailsAPP " + appDetailsAPP.Count);
                    Console.WriteLine("total app just installed wmiAppDetailsAPP " + wmiAppDetailsAPP.Count);
                    Console.WriteLine("total app just installed API " + appsByAppNameAPP.Count);
                    Console.WriteLine("total app just uninstalledApp " + uninstalledApp.Count);
                    Console.WriteLine("total app just myLog1034APP " + myLog1034APP.Count);
                    Console.WriteLine("total app just myLog11707APP " + myLog11707APP.Count);
                    Console.WriteLine("total app just appFromStoreAPP " + appFromStoreAPP.Count);


                    Console.WriteLine("total app just installed " + mergedInstalledApps.Count);

                    var result = WinApp.CompareAndCleanAppDictionaries(mergedInstalledApps, mergedUninstalledApps);
                    var updatedApps = result.UpdatedApps;
                    var removedApps = result.RemovedApps;
                    WinApp.ExportToExcel(updatedApps, removedApps);

                


                    break;

                case 4:

                // Get data from each function
                Dictionary<string, Tuple<string, string, string, string>> wmiAppDetails = WMI.GetAppByWMI();
                    Dictionary<string, Tuple<string, string>> registryAppDetails = MyRegistry.GetAppByRegistry("Software\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall");
                    Dictionary<string, Tuple<string, string>> registr2 = MyRegistry.GetAppByRegistry("Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall");
                    foreach (var pair in registr2)
                    {
                        if (!registryAppDetails.ContainsKey(pair.Key))
                        {
                            registryAppDetails[pair.Key] = pair.Value;
                        }
                    }
                    Dictionary<string, Tuple<string, string, string, string>> appsByAppName = WinApi.GetAppsUsingAPI();
                Dictionary<string, Tuple<string>> uninstalledApps = MyRegistry.UninstallApp();
                Dictionary<string, Tuple<DateTime, string, int>> myLog1034 = Helper.GetLog(1034);
                Dictionary<string, Tuple<DateTime, string, int>> myLog11707 = Helper.GetLog(11707);
                Dictionary<string, Tuple<DateTime, long>> appFromStore = Helper.CollectAppDetailsStore();


                Dictionary<string, Tuple<string, DateTime, long>> appDetails = new Dictionary<string, Tuple<string, DateTime, long>>();
                Helper.CollectAppDetails("C:\\Program Files", appDetails);
                Helper.CollectAppDetails("C:\\Program Files (x86)", appDetails);
                Helper.CollectAppDetails($"C:\\Users\\{Environment.UserName}\\AppData\\", appDetails);
                Helper.CollectAppDetails("C:\\Windows\\System32", appDetails);
                Helper.CollectAppDetails("C:\\ProgramData", appDetails);


                    
                // Create a new Excel package
                using (ExcelPackage excel = new ExcelPackage())
                {
                    // Add a worksheet for WMI app details
                    MyExcel.ExportToExcelSheet(wmiAppDetails, "WMI", excel, new string[] { "App Name", "Version", "Install Location", "Install State", "Install Date" });

                    // Add a worksheet for Registry app details
                    MyExcel.ExportToExcelSheet(registryAppDetails, "Registry", excel, new string[] { "App Name", "Version", "Install Location" });

                    MyExcel.ExportToExcelSheet(appsByAppName, "WIN API MSI", excel, new string[] { "App Name", "packageCode", "Install Location", "installDate", "installedSize" });

                    // Add a worksheet for Event Viewer apps Event ID 11707
                    MyExcel.ExportToExcelSheet(myLog11707, "Event Viewer Installed Apps", excel, new string[] { "App Name", "Time Generated", "Source", "Event ID" });

                    // Add a worksheet for store apps 
                    MyExcel.ExportToExcelSheet(appFromStore, "App Store", excel, new string[] { "App Name", "Time", "Size" });

                    // Add a worksheet for folder apps
                    MyExcel.ExportToExcelSheet(appDetails, "App Folder", excel, new string[] { "App Name", "dir", "Time", "Size" });

                    // Add a worksheet for Uninstalled apps
                    MyExcel.ExportToExcelSheet(uninstalledApps, "Uninstalled Apps", excel, new string[] { "App Name", "Status" });

                    // Add a worksheet for Event Viewer apps Event ID 1034
                    MyExcel.ExportToExcelSheet(myLog1034, "Event Viewer Uninstalled Apps", excel, new string[] { "App Name", "Time Generated", "Source", "Event ID" });



                    // Save to file
                    FileInfo excelFile = new FileInfo("AllAppDetails.xlsx");
                    try
                    {
                        excel.SaveAs(excelFile);
                        Console.WriteLine($"Excel file has been saved at: {excelFile.FullName}");
                    }
                    catch (IOException ioEx)
                    {
                        Console.WriteLine($"IO Error: {ioEx.Message}");
                    }
                    catch (UnauthorizedAccessException unAuthEx)
                    {
                        Console.WriteLine($"Access Error: You don't have permission to write to this location. {unAuthEx.Message}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"An unknown error occurred: {ex.Message}");
                    }
                    FileInfo existingFile = new FileInfo("AllAppDetails.xlsx");

                    Dictionary<string, int> appCount = new Dictionary<string, int>();
                    Dictionary<string, List<string>> appMethods = new Dictionary<string, List<string>>();

                    using (ExcelPackage package = new ExcelPackage(existingFile))
                    {
                        foreach (var worksheet in package.Workbook.Worksheets)
                        {
                            int rowCount = worksheet.Dimension.Rows;

                            for (int row = 2; row <= rowCount; row++) // Assuming the first row is header
                            {
                                string appName = worksheet.Cells[row, 1].Text;
                                string method = worksheet.Name; // Assuming the worksheet name is the method name

                                if (!string.IsNullOrEmpty(appName))
                                {
                                    if (appCount.ContainsKey(appName))
                                    {
                                        appCount[appName]++;
                                    }
                                    else
                                    {
                                        appCount[appName] = 1;
                                    }

                                    if (appMethods.ContainsKey(appName))
                                    {
                                        appMethods[appName].Add(method);
                                    }
                                    else
                                    {
                                        appMethods[appName] = new List<string>() { method };
                                    }
                                }
                            }
                        }

                        // Create a new worksheet for the summary
                        var summaryWorksheet = package.Workbook.Worksheets.Add("Summary");
                        summaryWorksheet.Cells[1, 1].Value = "App Name";
                        summaryWorksheet.Cells[1, 2].Value = "Count";
                        summaryWorksheet.Cells[1, 3].Value = "WMI";
                        summaryWorksheet.Cells[1, 4].Value = "WIN API MSI";
                        summaryWorksheet.Cells[1, 5].Value = "Registry";
                        summaryWorksheet.Cells[1, 6].Value = "Event Viewer Installed Apps";
                        summaryWorksheet.Cells[1, 7].Value = "App Store";
                        summaryWorksheet.Cells[1, 8].Value = "App Folder";
                        summaryWorksheet.Cells[1, 9].Value = "Uninstalled Apps";
                        summaryWorksheet.Cells[1, 10].Value = "Event Viewer Uninstalled Apps";



                        int summaryRow = 2;
                        foreach (var appName in appCount.Keys)
                        {
                            summaryWorksheet.Cells[summaryRow, 1].Value = appName;
                            summaryWorksheet.Cells[summaryRow, 2].Value = appCount[appName];

                            if (appMethods.ContainsKey(appName))
                            {
                                foreach (var method in appMethods[appName])
                                {
                                    switch (method)
                                    {
                                        case "WMI":
                                            summaryWorksheet.Cells[summaryRow, 3].Value = 1;
                                            break;
                                        case "WIN API MSI":
                                            summaryWorksheet.Cells[summaryRow, 4].Value = 1;
                                            break;
                                        case "Registry":
                                            summaryWorksheet.Cells[summaryRow, 5].Value = 1;
                                            break;
                                        case "Event Viewer Installed Apps":
                                            summaryWorksheet.Cells[summaryRow, 6].Value = 1;
                                            break;
                                        case "App Store":
                                            summaryWorksheet.Cells[summaryRow, 7].Value = 1;
                                            break;
                                        case "App Folder":
                                            summaryWorksheet.Cells[summaryRow, 8].Value = 1;
                                            break;
                                        case "Uninstalled Apps":
                                            summaryWorksheet.Cells[summaryRow, 9].Value = 1;
                                            break;
                                        case "Event Viewer Uninstalled Apps":
                                            summaryWorksheet.Cells[summaryRow, 10].Value = 1;
                                            break;
                                        default:
                                            break;
                                    }
                                }
                            }

                            summaryRow++;
                        }

                        package.Save();
                    }


                    Console.WriteLine("Summary sheet has been added to the Excel file.");

                    FileInfo existingFileForRegistry = new FileInfo("AllAppDetails.xlsx");
                    Dictionary<string, int> appCountFromRegistry = new Dictionary<string, int>();

                    using (ExcelPackage package = new ExcelPackage(existingFileForRegistry))
                    {
                        var registryWorksheet = package.Workbook.Worksheets["Registry"];
                        int rowCount = registryWorksheet.Dimension.Rows;

                        // Count apps that come from Registry
                        for (int row = 2; row <= rowCount; row++) // Assuming the first row is header
                        {
                            string appName = registryWorksheet.Cells[row, 1].Text;
                            if (!string.IsNullOrEmpty(appName))
                            {
                                appCountFromRegistry[appName] = 0; // Initialize
                            }
                        }

                        // Count these apps across all worksheets
                        foreach (var worksheet in package.Workbook.Worksheets)
                        {
                            rowCount = worksheet.Dimension.Rows;

                            for (int row = 2; row <= rowCount; row++) // Assuming the first row is header
                            {
                                string appName = worksheet.Cells[row, 1].Text;
                                if (appCountFromRegistry.ContainsKey(appName))
                                {
                                    appCountFromRegistry[appName]++;
                                }
                            }
                        }

                        // Create a new worksheet for the summary
                        var summaryWorksheet = package.Workbook.Worksheets.Add("Registry_Summary");
                        summaryWorksheet.Cells[1, 1].Value = "App Name";
                        summaryWorksheet.Cells[1, 2].Value = "Count";

                        int summaryRow = 2;
                        foreach (var kvp in appCountFromRegistry)
                        {
                            summaryWorksheet.Cells[summaryRow, 1].Value = kvp.Key;
                            summaryWorksheet.Cells[summaryRow, 2].Value = kvp.Value;
                            summaryRow++;
                        }

                        package.Save();
                    }

                    Console.WriteLine("Registry summary sheet has been added to the Excel file.");

                }

                break;

            case 5:
                Console.WriteLine("Exiting the program");
                continueMenu = false;
                break;

            default:
                Console.WriteLine("Invalid Choice");
                break;
        }
    } while (continueMenu);

    }
}










