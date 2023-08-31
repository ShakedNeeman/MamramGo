using System.Management;
using System;

public class WMIServiceDisplay
{
    public static void Display()
    {
        // create management scope object
        ManagementScope scope = new ManagementScope("\\\\.\\ROOT\\cimv2");

        //create object query
        ObjectQuery query = new ObjectQuery("SELECT * FROM Win32_Service");

        // create object searcher
        ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);

        // get a collection of WMI objects
        ManagementObjectCollection queryCollection = searcher.Get();

        // enumerate the collection.
        foreach (ManagementObject m in queryCollection)
        {
            // access properties of the WMI object
            Console.WriteLine($"Service Name : {m["Name"]}");
            Console.WriteLine($"Display Name : {m["DisplayName"]}");
            Console.WriteLine($"Status       : {m["State"]}");
            Console.WriteLine($"Start Mode: {m["StartMode"]}");
            Console.WriteLine($"Executable Path: {m["PathName"]}");
            Console.WriteLine("---------------------------------------");
        }
    }
}