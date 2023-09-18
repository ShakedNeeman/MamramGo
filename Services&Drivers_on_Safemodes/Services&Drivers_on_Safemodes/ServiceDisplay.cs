using System.Management;
using System;
using System.IO;
using OfficeOpenXml;

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
    public void SaveToWorksheet(ExcelWorksheet worksheet)
    {
        // Here we'll retrieve the data again and write it to the worksheet
        // This could be optimized by storing the results from the Display() method and reusing them, 
        // but for simplicity, I'm retrieving them again.
        int count = 0;
        int row = 1;

        worksheet.Cells[row, 1].Value = "Service Name";
        worksheet.Cells[row, 2].Value = "Path";
        worksheet.Cells[row, 3].Value = "File Name";

        ManagementScope scope = new ManagementScope("\\\\.\\ROOT\\cimv2");
        ObjectQuery query = new ObjectQuery("SELECT * FROM Win32_Service WHERE State = 'Running'");
        ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);
        ManagementObjectCollection queryCollection = searcher.Get();

        foreach (ManagementObject m in queryCollection)
        {
            string fullPath = m["PathName"]?.ToString();
            if (fullPath != null && fullPath.StartsWith("\""))
            {
                fullPath = fullPath.Split('\"')[1];
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

            row++;
            worksheet.Cells[row, 1].Value = m["Name"];
            worksheet.Cells[row, 2].Value = directory;
            worksheet.Cells[row, 3].Value = fileName;

            count++;
        }
    }
}
