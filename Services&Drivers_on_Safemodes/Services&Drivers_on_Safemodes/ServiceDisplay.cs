﻿using System.Management;
using System;
using System.IO;

public class ServiceDisplay
{
    public int Display()
    {
        int count = 0;

        // create management scope object
        ManagementScope scope = new ManagementScope("\\\\.\\ROOT\\cimv2");

        // create object query
        ObjectQuery query = new ObjectQuery("SELECT * FROM Win32_Service WHERE State = 'Running'");

        // create object searcher
        ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);

        // get a collection of WMI objects
        ManagementObjectCollection queryCollection = searcher.Get();

        // enumerate the collection.
        foreach (ManagementObject m in queryCollection)
        {
            // access properties of the WMI object
            string fullPath = m["PathName"]?.ToString();

            if (fullPath != null && fullPath.StartsWith("\""))
            {
                fullPath = fullPath.Split('\"')[1];  // This will get the path without quotes and without any arguments after the path.
            }

            string directory;
            string fileName;

            try
            {
                directory = fullPath != null ? Path.GetDirectoryName(fullPath) : "N/A";
                fileName = fullPath != null ? Path.GetFileName(fullPath) : "N/A";
            }
            catch (ArgumentException)
            {
                directory = "Error: Invalid path format";
                fileName = "Error: Invalid path format";
            }

            count++;
            Console.WriteLine($"Service Name : {m["Name"]}");
            Console.WriteLine($"Path         : {directory}");
            Console.WriteLine($"File Name    : {fileName}");
            Console.WriteLine("---------------------------------------");
        }
        return count;
    }
}
