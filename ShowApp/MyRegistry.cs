using Microsoft.Win32;
using System.Management;

namespace ShowApp
{
    internal class MyRegistry
    {

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

    }
}
