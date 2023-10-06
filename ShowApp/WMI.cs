using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading.Tasks;

namespace ShowApp
{
    internal class WMI
    {


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
                    string installLocation = m["InstallLocation"] as string;
                    if (string.IsNullOrEmpty(installLocation))
                    {
                        installLocation = m["InstallSource"] as string;
                    }
                    string installState = m["InstallState"]?.ToString() ?? "";
                    string installDateStr = m["InstallDate"] as string;
                    if (installDateStr != null)
                    {
                        //DateTime installDate = DateTime.ParseExact(installDateStr, "yyyyMMdd", null);
                        //app._date = installDate.ToString();
                        try
                        {
                            DateTime installDate = DateTime.ParseExact(installDateStr, "yyyyMMdd", null);
                            installDateStr = installDate.ToString();
                        }
                        catch (FormatException)
                        {
                            Console.WriteLine($"Failed to parse date: {installDateStr}");
                        }
                    }

                        // Add the details to the dictionary if the app name is not empty
                        if (!string.IsNullOrEmpty(appName))
                    {
                        appDetails[appName] = new Tuple<string, string, string, string>(appVersion, installLocation, installState, installDateStr);
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
    }
}
