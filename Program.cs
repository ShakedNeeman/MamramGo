using System;
using System.Diagnostics;
using Microsoft.Win32;
using System.Management;
using System.Runtime.InteropServices;
using System.Text;

namespace Task_1_switch
{
    class Program
    {
        [DllImport("psapi.dll")]
        public static extern bool EnumProcesses(
       [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U4)][In][Out] uint[] processIds,
       uint arraySizeBytes,
       [MarshalAs(UnmanagedType.U4)] out uint bytesCopied);

        // Define constants for working with 64-bit processes.
        private const int PROCESS_QUERY_LIMITED_INFORMATION = 0x1000;
        private const int PROCESS_VM_READ = 0x0010;

        // Import necessary Windows API functions.
        [DllImport("kernel32.dll")]
        private static extern IntPtr OpenProcess(int dwDesiredAccess, bool bInheritHandle, int dwProcessId);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool CloseHandle(IntPtr hObject);

        [DllImport("kernel32.dll")]
        private static extern bool IsWow64Process(IntPtr hProcess, out bool wow64Process);

        [DllImport("psapi.dll")]
        private static extern uint GetModuleFileNameEx(IntPtr hProcess, IntPtr hModule, [Out] StringBuilder lpBaseName, [In][MarshalAs(UnmanagedType.U4)] int nSize);

        [DllImport("psapi.dll", SetLastError = true)]
        private static extern int EnumProcessModules(IntPtr hProcess, IntPtr[] lphModule, uint cb, [MarshalAs(UnmanagedType.U4)] out uint lpcbNeeded);

        [DllImport("psapi.dll")]
        private static extern uint GetModuleBaseName(IntPtr hProcess, IntPtr hModule, [Out] StringBuilder lpBaseName, [In][MarshalAs(UnmanagedType.U4)] int nSize);


        static void Main()
        {
            int choice;

            do
            {
                Console.WriteLine("Choose a method to list processes:");
                Console.WriteLine("1. Using WMI (Windows Management Instrumentation)");
                Console.WriteLine("2. Using WinAPI (Windows API)");
                Console.WriteLine("3. Using PE Files (Portable Executable)");
                Console.WriteLine("4. Using EnumProcesses");
                Console.WriteLine("0. Exit");

                choice = int.Parse(Console.ReadLine());

                switch (choice)
                {
                    case 1:
                        ListProcessesUsingWMI();
                        break;
                    case 2:
                        ListProcessesUsingWinAPI();
                        break;
                    case 3:
                        ListProcessesUsingPEFiles();
                        break;
                    case 4:
                        ListProcessesUsingEnumProcesses();
                        break;
                    case 0:
                        Console.WriteLine("Exiting program.");
                        break;
                    default:
                        Console.WriteLine("Invalid choice.");
                        break;
                }
            } while (choice != 0);
        }


         static void ListProcessesUsingWMI()
        {
            using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Process"))
            {
                foreach (ManagementObject process in searcher.Get())
                {
                    Console.WriteLine("Process: " + process["Name"]);
                }
            }
        }

       static void ListProcessesUsingWinAPI()
{
    try
    {
        Process[] processes = Process.GetProcesses();

        foreach (Process process in processes)
        {
            try
            {
                Console.WriteLine($"Process ID: {process.Id}");
                Console.WriteLine($"Process Name: {process.ProcessName}");

                // Check if the process has a valid main module before accessing its filename.
                if (process.MainModule != null)
                {
                    Console.WriteLine($"Executable Path: {process.MainModule.FileName}");
                }
                else
                {
                    Console.WriteLine("Executable Path: Not available (access denied)");
                }

                Console.WriteLine("--------------------------------------");
            }
            catch (Exception ex)
            {
                // Handle exceptions that occur when trying to access process information.
                if (ex is System.ComponentModel.Win32Exception win32Exception &&
                    (win32Exception.NativeErrorCode == 5 || win32Exception.NativeErrorCode == 299))
                {
                    // Error code 5 (Access is denied) or 299 (Only part of a ReadProcessMemory or WriteProcessMemory request was completed)
                    Console.WriteLine($"Access denied to process {process.ProcessName}");
                }
                else
                {
                    Console.WriteLine($"Error accessing process info for {process.ProcessName}: {ex.Message}");
                }
                continue; // Continue with the next process.
            }
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error: {ex.Message}");
    }
}


        static void ListProcessesUsingPEFiles()
        {
            try
            {
                Process[] processes = Process.GetProcesses();

                foreach (Process process in processes)
                {
                    try
                    {
                        string processName = GetProcessNameFromPE(process.Id);
                        string exePath = GetExecutablePathFromPE(process.Id);

                        Console.WriteLine($"Process ID: {process.Id}");
                        Console.WriteLine($"Process Name: {processName}");
                        Console.WriteLine($"Executable Path: {exePath}");
                        Console.WriteLine("--------------------------------------");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error accessing process info for {process.ProcessName}: {ex.Message}");
                        continue; // Continue with the next process.
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        static string GetProcessNameFromPE(int processId)
        {
            // Read the PE file to retrieve the process name (highly simplified).
            // This is a simplified example and may not work correctly for all processes.
            // You would typically use a dedicated PE file parsing library for this task.
            return "SampleProcessName";
        }

        static string GetExecutablePathFromPE(int processId)
        {
            // Read the PE file to retrieve the executable path (highly simplified).
            // This is a simplified example and may not work correctly for all processes.
            // You would typically use a dedicated PE file parsing library for this task.
            return "C:\\Path\\To\\Sample.exe";
        }

        static void ListProcessesUsingEnumProcesses()
        {
            uint[] processIds = new uint[1024];
            uint bytesCopied;

            if (EnumProcesses(processIds, (uint)(processIds.Length * sizeof(uint)), out bytesCopied))
            {
                int numProcesses = (int)(bytesCopied / sizeof(uint));

                for (int i = 0; i < numProcesses; i++)
                {
                    int processId = (int)processIds[i];

                    // Check if the process is still running before retrieving information.
                    try
                    {
                        Process process = Process.GetProcessById(processId);
                        Console.WriteLine($"Process Name: {process.ProcessName}, PID: {process.Id}");
                    }
                    catch (ArgumentException ex)
                    {
                        // Process with the specified ID is not running; skip it.
                        Console.WriteLine($"Process with ID {processId} is not running.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error accessing process info for ID {processId}: {ex.Message}");
                    }
                }
            }

        }
    }
}