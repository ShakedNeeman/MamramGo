using Microsoft.Win32;
using System.Management;
using OfficeOpenXml;
using System.Diagnostics;
using ShowApp;

bool continueMenu = true;
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


do
{
    Console.WriteLine("Enter a number to select a display operation method");
    Console.WriteLine("1-Show All The App By WMI");
    Console.WriteLine("2-Show All The App By Registry");
    Console.WriteLine("3-Show All The App By API");
    Console.WriteLine("4-Export To CSV ");



    Console.WriteLine("5-EXIT");

    int Number = int.Parse(Console.ReadLine());


    switch (Number)
    {
        case 1:
            WMI.GetAppByWMI(); // Assuming you have GetAppByWMI() defined as before                  
            break;
        case 2:
            MyRegistry.GetAppByRegistry();
            break;
        case 3:
            WinApi.GetAppsUsingAPI();
            break;

        case 4:

            // Get data from each function
            Dictionary<string, Tuple<string, string, string, string>> wmiAppDetails = WMI.GetAppByWMI();
            Dictionary<string, (string Version, string InstallLocation)> registryAppDetails = MyRegistry.GetAppByRegistry();
            Dictionary<string, Tuple<string, string, string, string>> appsByAppName = WinApi.GetAppsUsingAPI();
            Dictionary<string, string> uninstalledApps = MyRegistry.UninstallApp();
            Dictionary<string, Tuple<DateTime, string, int>> myLog1034 = Helper.GetLog(1034);
            Dictionary<string, Tuple<DateTime, string, int>> myLog11707 = Helper.GetLog(11707);


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

                // Add a worksheet for Uninstalled apps
                MyExcel.ExportToExcelSheet(uninstalledApps, "Uninstalled Apps", excel, new string[] { "App Name", "Status" });

                // Add a worksheet for Event Viewer apps Event ID 1034
                MyExcel.ExportToExcelSheet(myLog1034, "Event Viewer Uninstalled Apps", excel, new string[] { "App Name", "Time Generated", "Source", "Event ID" });

                // Add a worksheet for Event Viewer apps Event ID 11707
                MyExcel.ExportToExcelSheet(myLog11707, "Event Viewer Installed Apps", excel, new string[] { "App Name", "Time Generated", "Source", "Event ID" });


                MyExcel.ExportToExcelSheet(appDetails, "App Folder", excel, new string[] { "App Name", "dir", "Time", "size" });


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
                    summaryWorksheet.Cells[1, 6].Value = "Uninstalled Apps";
                    summaryWorksheet.Cells[1, 7].Value = "Event Viewer Uninstalled Apps";
                    summaryWorksheet.Cells[1, 8].Value = "Event Viewer Installed Apps";
                    summaryWorksheet.Cells[1, 9].Value = "App Folder";



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
                                    case "Uninstalled Apps":
                                        summaryWorksheet.Cells[summaryRow, 6].Value = 1;
                                        break;
                                    case "Event Viewer Uninstalled Apps":
                                        summaryWorksheet.Cells[summaryRow, 7].Value = 1;
                                        break;
                                    case "Event Viewer Installed Apps":
                                        summaryWorksheet.Cells[summaryRow, 8].Value = 1;
                                        break;
                                    case "App Folder":
                                        summaryWorksheet.Cells[summaryRow, 9].Value = 1;
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





  







