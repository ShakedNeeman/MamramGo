using Microsoft.Win32; 
using System; 

public class RegistryServiceDisplay
{
    public static void Display()
    {
        // Open the Registry key where Windows services are listed.
        using (RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Services"))
        {
            // Check if the key exists.
            if (key != null)
            {
                // Loop through all subkeys, which correspond to the service names.
                foreach (string serviceName in key.GetSubKeyNames())
                {
                    // Open each service's individual Registry key.
                    using (RegistryKey serviceKey = key.OpenSubKey(serviceName))
                    {
                        // Check if the service key is valid.
                        if (serviceKey != null)
                        {
                            // Get the display name and start type of the service.
                            object displayNameObj = serviceKey.GetValue("DisplayName");
                            object startTypeObj = serviceKey.GetValue("Start");
                            object serviceTypeObj = serviceKey.GetValue("Type");
                            object errorControlObj = serviceKey.GetValue("ErrorControl");


                            // The "Start" value is used as a filter to identify actual services. 
                            // Some entries in the "SYSTEM\CurrentControlSet\Services" subkey are not actual services,
                            // but other types of drivers or settings. These usually do not have a "Start" value,
                            // so they can be filtered out.
                            if (startTypeObj != null)
                            {
                                // If there are null values, display "N/A"
                                Console.WriteLine($"Service Name    : {serviceName}");
                                Console.WriteLine($"Display Name    : {displayNameObj ?? "N/A"}"); 
                                Console.WriteLine($"Start Type      : {startTypeObj ?? "N/A"}");
                                Console.WriteLine($"Service Type    : {serviceTypeObj ?? "N/A"}");
                                Console.WriteLine($"Error Control   : {errorControlObj ?? "N/A"}");
                                Console.WriteLine("----------------------------------------"); // Separator for better readability
                            }

                           
                           
                        }
                    }
                }
            }
        }
    }
}
