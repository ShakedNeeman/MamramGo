using System;
using System.Management;
using System.IO;

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
    }
}