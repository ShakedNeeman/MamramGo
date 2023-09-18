using System;
using System.Management;
using System.IO;
using OfficeOpenXml;

public class Program
{
    public static void Main()
    {
        Console.WriteLine("Displaying Running Services...");
        Console.WriteLine("===============================================");
        ServiceDisplay serviceDisplay = new ServiceDisplay();
        int serviceCount = serviceDisplay.Display();

        Console.WriteLine("\n\nDisplaying Running System Drivers...");
        Console.WriteLine("===============================================");
        SystemDriverDisplay driverDisplay = new SystemDriverDisplay();
        int driverCount = driverDisplay.Display();

        Console.WriteLine($"\nTotal Running Services: {serviceCount}");
        Console.WriteLine($"Total Running System Drivers: {driverCount}");

        SaveToExcel(serviceDisplay, driverDisplay);
    }

    public static void SaveToExcel(ServiceDisplay serviceDisplay, SystemDriverDisplay driverDisplay)
    {
        // Initialize a new Excel package (spreadsheet)
        using (ExcelPackage package = new ExcelPackage())
        {
            // Create a new worksheet for services in the Excel file
            ExcelWorksheet servicesWorksheet = package.Workbook.Worksheets.Add("Services");
            // Write the service data to the newly created services worksheet
            serviceDisplay.SaveToWorksheet(servicesWorksheet);

            // Create a new worksheet for system drivers in the Excel file
            ExcelWorksheet driversWorksheet = package.Workbook.Worksheets.Add("Drivers");
            // Write the system driver data to the newly created drivers worksheet
            driverDisplay.SaveToWorksheet(driversWorksheet);

            // Define the path where the Excel file will be saved (on the desktop)
            string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "ServicesAndDrivers.xlsx");
            // Save the Excel package (spreadsheet) to the defined path
            package.SaveAs(new FileInfo(path));

            // Notify the user where the Excel file has been saved
            Console.WriteLine($"\nSaved to {path}");
        }
    }
}

