using Microsoft.Win32;
using System.Management;
using OfficeOpenXml;
using System.Diagnostics;
using ShowApp;

class Program
{
    static void Main(string[] args)
    {
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
                    Dictionary<string, Tuple<EventLogEntryType, DateTime, string, int>> myLog = Helper.GetLog1034();

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
                        MyExcel.ExportToExcelSheet(wmiAppDetails, "WMI_AppDetails", excel, new string[] { "App Name", "Version", "Install Location", "Install State", "Install Date" });

                        // Add a worksheet for Registry app details
                        MyExcel.ExportToExcelSheet(registryAppDetails, "Registry_AppDetails", excel, new string[] { "App Name", "Version", "Install Location" });

                        MyExcel.ExportToExcelSheet(appsByAppName, "API_AppDetails", excel, new string[] { "App Name", "packageCode", "Install Location", "installDate", "installedSize" });

                        // Add a worksheet for Uninstalled apps
                        MyExcel.ExportToExcelSheet(uninstalledApps, "Uninstalled_Apps", excel, new string[] { "App Name", "Status" });

                        // Add a worksheet for Event Viewer apps
                        MyExcel.ExportToExcelSheet(myLog, "Event Viewer Log Application", excel, new string[] { "App Name", "Entry Type", "Time Generated", "Source", "Event ID" });

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

                        using (ExcelPackage package = new ExcelPackage(existingFile))
                        {
                            foreach (var worksheet in package.Workbook.Worksheets)
                            {
                                int rowCount = worksheet.Dimension.Rows;

                                for (int row = 2; row <= rowCount; row++) // Assuming the first row is header
                                {
                                    string appName = worksheet.Cells[row, 1].Text;

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
                                    }
                                }
                            }

                            // Create a new worksheet for the summary
                            var summaryWorksheet = package.Workbook.Worksheets.Add("Summary");
                            summaryWorksheet.Cells[1, 1].Value = "App Name";
                            summaryWorksheet.Cells[1, 2].Value = "Count";

                            int summaryRow = 2;
                            foreach (var kvp in appCount)
                            {
                                summaryWorksheet.Cells[summaryRow, 1].Value = kvp.Key;
                                summaryWorksheet.Cells[summaryRow, 2].Value = kvp.Value;
                                summaryRow++;
                            }

                            package.Save();
                        }

                        Console.WriteLine("Summary sheet has been added to the Excel file.");

                        FileInfo existingFileForRegistry = new FileInfo("AllAppDetails.xlsx");
                        Dictionary<string, int> appCountFromRegistry = new Dictionary<string, int>();

                        using (ExcelPackage package = new ExcelPackage(existingFileForRegistry))
                        {
                            var registryWorksheet = package.Workbook.Worksheets["Registry_AppDetails"];
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





  







