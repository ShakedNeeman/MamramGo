using Microsoft.Win32;
using System.Diagnostics;
using System.Management;
using System.Collections.Generic; // Required for Dictionary class


class Program
    {
        static void Main(string[] args)
        {
            bool continueMenu = true;

            do
            {
                Console.WriteLine("Enter a number to select a display operation method");
                Console.WriteLine("1-Show All The App By WMI");
                Console.WriteLine("2-Show All The App By Registry");
                Console.WriteLine("3-Show All The App By PowerShell");
                Console.WriteLine("4-Show All The App By CMD");
                Console.WriteLine("5-Show All The App By API");


            Console.WriteLine("6-EXIT");

                int dayNumber = int.Parse(Console.ReadLine());


                switch (dayNumber)
                {
                    case 1:
                        GetAppByWMI();
                        break;
                    case 2:
                        GetAppByRegistry();
                        break;
                    case 3:
                        GetAppsByPowerShell();
                        break;
                    case 4:
                        GetAppsUsingCmd();
                        break;
                    case 5:
                        WinApi.GetAppsUsingAPI();
                    break;
                case 6:
                        Console.WriteLine("Exiting the program");
                        continueMenu = false;
                        break;
                    default:
                        Console.WriteLine("Invalid Choice");
                        break;
                }
            } while (continueMenu);
        }


    public static void GetAppByWMI()
    {
        try
        {
            // Initialize a dictionary to store unique application details
            Dictionary<string, string> appDetails = new Dictionary<string, string>();

            // Create an object of WMI for querying Win32_Product class
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Product");

            // Execute the WMI query and get the result
            ManagementObjectCollection queryCollection = searcher.Get();
            // Initialize message and counter
            Console.WriteLine("Result Get-WmiObject -Class Win32_Product | Format-List:");

            // Iterate through the WMI query result
            foreach (ManagementObject m in queryCollection)
            {
                // Extract the 'Name' and 'Version' from the current WMI object
                string appName = m["Name"]?.ToString() ?? "";
                string appVersion = m["Version"]?.ToString() ?? "";

                // Check if both Name and Version are not empty, to consider it as an installed application
                if (!string.IsNullOrEmpty(appName) && !string.IsNullOrEmpty(appVersion))
                {
                    // Check if this application name already exists in the dictionary
                    if (!appDetails.ContainsKey(appName))
                    {
                        // Add the application details to the dictionary
                        appDetails[appName] = appVersion;
                    }
                }
            }

            // Display the details from the dictionary
            foreach (var detail in appDetails)
            {
                Console.WriteLine("-----------------------------------");
                Console.WriteLine($"Application Name: {detail.Key}");
                Console.WriteLine($"Version: {detail.Value}");
            }

            // Output the total number of unique applications found
            Console.WriteLine($"Total Unique Apps = {appDetails.Count}");
        }
        catch (ManagementException e)
        {
            // Display an error message if a WMI-specific exception occurs
            Console.WriteLine("ERROR WMI: " + e.Message);
        }
    }

    public static void GetAppByRegistry()
    {
        try
        {
            // Initialize a dictionary to store application details
            Dictionary<string, (string Version, string InstallLocation)> appDetails = new Dictionary<string, (string, string)>();

            // Open the Registry key for installed applications
            RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall");

            // Output header message
            Console.WriteLine("Installed Applications Information:");

            // Loop through each subkey (application) under the "Uninstall" key
            foreach (string subKeyName in key.GetSubKeyNames())
            {
                // Open the subkey (application key)
                RegistryKey subKey = key.OpenSubKey(subKeyName);

                // Check for a null key just to be safe
                if (subKey != null)
                {
                    // Retrieve the application's display name, version, and install location from the subkey
                    object appName = subKey.GetValue("DisplayName");
                    object appVersion = subKey.GetValue("DisplayVersion");
                    object installLocation = subKey.GetValue("InstallLocation");

                    // Check if the application name is not null (it's one of the main properties)
                    if (appName != null)
                    {
                        // Check if this application name already exists in the dictionary
                        if (!appDetails.ContainsKey(appName.ToString()))
                        {
                            // Add the application to the dictionary
                            appDetails[appName.ToString()] = (appVersion?.ToString() ?? "N/A", installLocation?.ToString() ?? "N/A");
                        }
                    }
                }
            }

            // Display the details from the dictionary
            foreach (var detail in appDetails)
            {
                Console.WriteLine("-----------------------------------");
                Console.WriteLine($"Application Name: {detail.Key}");
                Console.WriteLine($"Version: {detail.Value.Version}");
                Console.WriteLine($"Install Location: {detail.Value.InstallLocation}");
            }

            // Output the total number of unique applications found
            Console.WriteLine($"Total Unique Apps = {appDetails.Count}");
        }
        catch (Exception e)
        {
            // Output any exceptions that were thrown
            Console.WriteLine("An error occurred: " + e.Message);
        }
    }

    public static void GetAppsByPowerShell()
        {

            // Create a ProcessStartInfo object to configure the PowerShell process
            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = "powershell", // Set the executable to PowerShell
                RedirectStandardOutput = true, // Redirect the standard output
                RedirectStandardError = true, // Redirect the standard error
                RedirectStandardInput = true, // Redirect the standard input
                UseShellExecute = false, // Don't use the system shell to execute the process
                CreateNoWindow = true // Don't create a window for the process
            };

            // Create a Process object and assign the ProcessStartInfo
            Process process = new Process
            {
                StartInfo = psi
            };

            // Start the PowerShell process
            process.Start();

            // Execute PowerShell command to get application information
            string command = "Get-WmiObject -Class Win32_Product | Select-Object Name, Version, InstallLocation";

            // Write the command to the standard input of the process
            process.StandardInput.WriteLine(command);

            // Close the standard input to indicate no more input will be written
            process.StandardInput.Close();

            // Read the output and error streams
            string output = process.StandardOutput.ReadToEnd();
            string error = process.StandardError.ReadToEnd();

            // Wait for the process to exit
            process.WaitForExit();

            // Check if there's any error and display it
            if (!string.IsNullOrWhiteSpace(error))
            {
                Console.WriteLine("PowerShell Error: " + error);
            }
            else
            {
                // Display the installed applications information
                Console.WriteLine("Installed Applications Information:");
                Console.WriteLine(output);
            // Split the output by new lines
            string[] lines = output.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);

            // Count the number of lines, each line represents one application
            int numberOfApps = lines.Length;

            Console.WriteLine($"Total Apps = {numberOfApps}");
        }
        }
    static public void GetAppsUsingCmd()
    {
        try
        {
            // Create a new process object
            Process process = new Process();

            // Define process start information
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                // Hide the cmd window
                WindowStyle = ProcessWindowStyle.Hidden,

                // Set process executable to cmd.exe
                FileName = "cmd.exe",

                // Redirect standard output and input to manipulate the stream
                RedirectStandardOutput = true,
                RedirectStandardInput = true,

                // Disable shell execute to allow redirection
                UseShellExecute = false,

                // Do not create a new cmd window
                CreateNoWindow = true
            };

            // Assign the configured start information to the process
            process.StartInfo = startInfo;

            // Start the process
            process.Start();

            // Define the WMIC command to list installed applications
            string cmdCommand = "wmic product get name, version, PackageCache";

            // Write the WMIC command to the process stdin stream
            process.StandardInput.WriteLine(cmdCommand);

            // Write 'exit' to process stdin stream to close it
            process.StandardInput.WriteLine("exit");

            // Read the entire standard output stream of the process
            string output = process.StandardOutput.ReadToEnd();

            // Wait for the process to exit
            process.WaitForExit();

            // Output the results
            Console.WriteLine("WMIC Output:");
            Console.WriteLine(output);
        }
        catch (Exception ex)
        {
           
            // General exception handler
            Console.WriteLine("An error occurred in GetAppsUsingCmd: " + ex.Message);
        }
    }
}



            

    


 
    



