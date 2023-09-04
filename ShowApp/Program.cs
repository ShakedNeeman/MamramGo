using Microsoft.Win32;
using System.Management;
using OfficeOpenXml;



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
                        GetAppByWMI(); // Assuming you have GetAppByWMI() defined as before                  
                    break;
                    case 2:
                        GetAppByRegistry();              
                    break;
                    case 3:
                        WinApi.GetAppsUsingAPI();
                    break;

                case 4:

                    // Get data from each function
                    Dictionary<string, Tuple<string, string, string, string>> wmiAppDetails = GetAppByWMI();
                    Dictionary<string, (string Version, string InstallLocation)> registryAppDetails = GetAppByRegistry();
                    Dictionary<string, Tuple<string, string, string, string>> appsByAppName = WinApi.GetAppsUsingAPI();
                    Dictionary<string, string> uninstalledApps = UninstallApp();

                    // Create a new Excel package
                    using (ExcelPackage excel = new ExcelPackage())
                    {
                        // Add a worksheet for WMI app details
                        ExportToExcelSheet(wmiAppDetails, "WMI_AppDetails", excel, new string[] { "App Name", "Version", "Install Location", "Install State", "Install Date" });

                        // Add a worksheet for Registry app details
                        ExportToExcelSheet(registryAppDetails, "Registry_AppDetails", excel, new string[] { "App Name", "Version", "Install Location" });

                        // Add a worksheet for Uninstalled apps
                        ExportToExcelSheet(uninstalledApps, "Uninstalled_Apps", excel, new string[] { "App Name", "Status" });

                        ExportToExcelSheet(appsByAppName, "API_AppDetails", excel, new string[] { "App Name", "packageCode", "Install Location", "installDate", "installedSize" });

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


    public static Dictionary<string, Tuple<string, string, string, string>> GetAppByWMI()
    {
        // Initialize a dictionary to hold application details
        Dictionary<string, Tuple<string, string, string, string>> appDetails = new Dictionary<string, Tuple<string, string, string, string>>();

        try
        {
            // Query WMI for Win32_Product entries
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Product");
            ManagementObjectCollection queryCollection = searcher.Get();

            // Loop through all entries to get app details
            foreach (ManagementObject m in queryCollection)
            {
                // Retrieve app details from WMI properties
                string appName = m["Name"]?.ToString() ?? "";
                string appVersion = m["Version"]?.ToString() ?? "";
                string installLocation = m["InstallLocation"]?.ToString() ?? "Unknown";
                string installState = m["InstallState"]?.ToString() ?? "";
                string installDate = m["InstallDate"]?.ToString() ?? "";

                // Add the details to the dictionary if the app name is not empty
                if (!string.IsNullOrEmpty(appName))
                {
                    appDetails[appName] = new Tuple<string, string, string, string>(appVersion, installLocation, installState, installDate);
                }
            }

            // Print app details to the console (for debug or logging)
            int index = 1;
            //foreach (var detail in appDetails)
            //{
            //    Console.WriteLine("-----------------------------------");
            //    Console.WriteLine($"{index}");
            //    Console.WriteLine($"Application Name: {detail.Key}");
            //    Console.WriteLine($"Version: {detail.Value.Item1}");
            //    Console.WriteLine($"Install Location: {detail.Value.Item2}");
            //    Console.WriteLine($"Install State: {detail.Value.Item3}");
            //    Console.WriteLine($"Install Date: {detail.Value.Item4}");
            //    index++;
            //}

            //// Print the total number of unique apps found
            //Console.WriteLine($"Total Unique Apps = {appDetails.Count}");
        }
        catch (ManagementException e)
        {
            // Log any exceptions that occur during the WMI query
            Console.WriteLine("An error occurred while querying for WMI data: " + e.Message);
        }

        // Return the dictionary containing app details
        return appDetails;
    }

    // Fetch installed apps from Registry and return them as a dictionary
    public static Dictionary<string, (string Version, string InstallLocation)> GetAppByRegistry()
    {
        // Initialize a dictionary to hold application details
        Dictionary<string, (string Version, string InstallLocation)> appDetails = new Dictionary<string, (string, string)>();

        try
        {
            // Open the Registry key where installed apps are listed
            RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall");

            // Loop through all subkeys to get app details
            foreach (string subKeyName in key.GetSubKeyNames())
            {
                // Open each subkey to read its values
                RegistryKey subKey = key.OpenSubKey(subKeyName);

                // Check for null and proceed
                if (subKey != null)
                {
                    // Read app details from Registry values
                    object appName = subKey.GetValue("DisplayName");
                    object appVersion = subKey.GetValue("DisplayVersion");
                    object installLocation = subKey.GetValue("InstallLocation");

                    // Add the details to the dictionary if app name exists
                    if (appName != null)
                    {
                        string appVersionString = appVersion?.ToString();
                        string installLocationString = installLocation?.ToString();

                        if (appVersionString != null || installLocationString != null)
                        {
                            appDetails[appName.ToString()] = (appVersionString, installLocationString);
                        }
                    }
                }
            }
        }
        catch (Exception e)
        {
            // Log any exceptions that occur
            Console.WriteLine("An error occurred: " + e.Message);
        }

        // Return the dictionary containing app details
        return appDetails;
    }



    public static Dictionary<string, string> UninstallApp()
    {
        // Initialize a dictionary to hold uninstalled application details
        Dictionary<string, string> uninstalledApps = new Dictionary<string, string>();

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
                object appName = subKey.GetValue("DisplayName");
                object uninstallInfo = subKey.GetValue("UninstallString");

                // Add the details to the dictionary if app name exists but uninstall info doesn't
                if (appName != null && uninstallInfo == null)
                {
                    uninstalledApps[appName.ToString()] = "Uninstalled";
                }
            }
        }

        // Return the dictionary containing uninstalled app details
        return uninstalledApps;
    }

    static void ExportToExcelSheet<T>(Dictionary<string, T> data, string sheetName, ExcelPackage excel, string[] headers)
    {
        var worksheet = excel.Workbook.Worksheets.Add(sheetName);

        // Add headers to the worksheet
        for (int i = 0; i < headers.Length; i++)
        {
            worksheet.Cells[1, i + 1].Value = headers[i];
        }

        // Initialize the row counter
        int row = 2;

        // Loop through each entry in the dictionary
        foreach (var entry in data)
        {
            // Add the key to the first column of the current row
            worksheet.Cells[row, 1].Value = entry.Key;

            // Check the type of the value
            if (entry.Value is Tuple<string, string, string, string> tupleValue)
            {
                // Manually unpack the tuple and add the items to the worksheet
                worksheet.Cells[row, 2].Value = tupleValue.Item1;
                worksheet.Cells[row, 3].Value = tupleValue.Item2;
                worksheet.Cells[row, 4].Value = tupleValue.Item3;
                worksheet.Cells[row, 5].Value = tupleValue.Item4;
            }
            else if (entry.Value is ValueTuple<string, string> valueTuple)
            {
                // Manually unpack the value tuple and add the items to the worksheet
                worksheet.Cells[row, 2].Value = valueTuple.Item1;
                worksheet.Cells[row, 3].Value = valueTuple.Item2;
            }
            else
            {
                // Convert the value to a string and add it to the second column of the current row
                worksheet.Cells[row, 2].Value = entry.Value.ToString();
            }

            // Increment the row counter
            row++;
        }
    }

}













