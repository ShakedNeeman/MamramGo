﻿using Microsoft.Win32;
using System.Globalization;
using System.Management;

namespace ShowApp
{
    internal class MyRegistry
    {

        // Fetch installed apps from Registry and return them as a dictionary
        public static Dictionary<string, Tuple<string, string>> GetAppByRegistry(string path)
        {
            // Initialize a dictionary to hold application details
            Dictionary<string, Tuple<string, string>> appDetails = new Dictionary<string, Tuple<string, string>>();

            try
            {
                // Open the Registry key where installed apps are listed
                RegistryKey key = Registry.LocalMachine.OpenSubKey(path);

                // Loop through all subkeys to get app details
                foreach (string subKeyName in key.GetSubKeyNames())
                {
                    // Open each subkey to read its values
                    RegistryKey subKey = key.OpenSubKey(subKeyName);

                    // Check for null and proceed
                    if (subKey != null)
                    {
                        // Read app details from Registry values
                        string appName = subKey.GetValue("DisplayName") as string;
                        string appVersion = subKey.GetValue("DisplayVersion") as string;
                        string installLocation = subKey.GetValue("InstallLocation") as string;
                        if (string.IsNullOrEmpty(installLocation))
                        {
                            installLocation = subKey.GetValue("InstallSource") as string;
                        }

                        // Add the details to the dictionary if app name exists
                        if (appName != null)
                        {
                            string appVersionString = appVersion;
                            string installLocationString = installLocation;

                            if (appVersionString != null || installLocationString != null)
                            {
                                appDetails[appName.ToString()] = new Tuple<string, string>(appVersionString, installLocationString);
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



        public static Dictionary<string, Tuple<string>> UninstallApp()
        {
            // Initialize a dictionary to hold uninstalled application details
            Dictionary<string, Tuple<string>> uninstalledApps = new Dictionary<string, Tuple<string>>();

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

                  
                }
            }

            // Return the dictionary containing uninstalled app details
            return uninstalledApps;
        }






    }
    }
