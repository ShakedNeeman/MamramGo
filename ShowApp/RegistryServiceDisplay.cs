using Microsoft.Win32;
using System;

public class RegistryServiceDisplay
{
    public static void Display()
    {
        // Connect to the registry and access the services subkey.
        using (RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Services"))
        {
            if (key != null)
            {
                // Print a header for better clarity.
                Console.WriteLine($"{"Service Name",-50} | {"Display Name",-100}");
                Console.WriteLine(new string('-', 155));  // Separator line.

                // Iterate through each service subkey.
                foreach (string serviceName in key.GetSubKeyNames())
                {
                    using (RegistryKey serviceKey = key.OpenSubKey(serviceName))
                    {
                        if (serviceKey != null)
                        {
                            object displayNameObj = serviceKey.GetValue("DisplayName");

                            // Check if display name exists and doesn't start with '@'.
                            if (displayNameObj != null && !displayNameObj.ToString().StartsWith("@"))
                            {
                                Console.WriteLine($"{serviceName,-50} | {displayNameObj,-100}");
                            }
                        }
                    }
                }
            }
        }
    }
}
