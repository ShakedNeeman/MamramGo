using System;
using System.Runtime.InteropServices;
using System.Text;

public class WinAPIServiceDisplay
{

    [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern IntPtr OpenSCManager(string machineName, string databaseName, uint dwDesiredAccess);

    [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern bool EnumServicesStatusEx(IntPtr hSCManager, int InfoLevel, uint dwServiceType, uint dwServiceState, IntPtr lpServices, uint cbBufSize, out uint pcbBytesNeeded, out uint lpServicesReturned, ref uint lpResumeHandle, string pszGroupName);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct ENUM_SERVICE_STATUS_PROCESS
    {
        public IntPtr lpServiceName;
        public IntPtr lpDisplayName;
        public SERVICE_STATUS_PROCESS ServiceStatusProcess;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct SERVICE_STATUS_PROCESS
    {
        public uint dwServiceType;
        public uint dwCurrentState;
        public uint dwControlsAccepted;
        public uint dwWin32ExitCode;
        public uint dwServiceSpecificExitCode;
        public uint dwCheckPoint;
        public uint dwWaitHint;
        public uint dwProcessId;
        public uint dwServiceFlags;
    }

    const uint SC_MANAGER_ENUMERATE_SERVICE = 4;
    const uint SERVICE_WIN32 = 0x00000030;
    const uint SERVICE_STATE_ALL = 0x00000003;

    public static void Display()
    {
        IntPtr scm = OpenSCManager(null, null, SC_MANAGER_ENUMERATE_SERVICE);

        if (scm == IntPtr.Zero)
        {
            Console.WriteLine("Failed to open service control manager.");
            return;
        }

        uint bytesNeeded = 0;
        uint servicesReturned = 0;
        uint resumeHandle = 0;
        uint bufSize = 0;

        // First call to get the size needed
        EnumServicesStatusEx(scm, 0, SERVICE_WIN32, SERVICE_STATE_ALL, IntPtr.Zero, bufSize, out bytesNeeded, out servicesReturned, ref resumeHandle, null);
        bufSize = bytesNeeded;

        IntPtr buffer = Marshal.AllocHGlobal((int)bufSize);

        // Actual call to list services
        if (EnumServicesStatusEx(scm, 0, SERVICE_WIN32, SERVICE_STATE_ALL, buffer, bufSize, out bytesNeeded, out servicesReturned, ref resumeHandle, null))
        {
            for (int i = 0; i < servicesReturned; i++)
            {
                IntPtr ithService = new IntPtr((long)buffer + i * Marshal.SizeOf(typeof(ENUM_SERVICE_STATUS_PROCESS)));
                ENUM_SERVICE_STATUS_PROCESS serviceStatus = (ENUM_SERVICE_STATUS_PROCESS)Marshal.PtrToStructure(ithService, typeof(ENUM_SERVICE_STATUS_PROCESS));

                Console.WriteLine($"Service Name: {Marshal.PtrToStringUni(serviceStatus.lpServiceName)}");
                Console.WriteLine($"Display Name: {Marshal.PtrToStringUni(serviceStatus.lpDisplayName)}");
                Console.WriteLine($"Current State: {serviceStatus.ServiceStatusProcess.dwCurrentState}");
                Console.WriteLine("-----------------------------------------------------");
            }
        }

        Marshal.FreeHGlobal(buffer);
    }

}
