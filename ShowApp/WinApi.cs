using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;

class WinApi
{
    [DllImport("msi.dll", CharSet = CharSet.Unicode)]
    static extern int MsiEnumProducts(int iProductIndex, StringBuilder lpProductBuf);

    [DllImport("msi.dll", SetLastError = true)]
    static extern Int32 MsiGetProductInfo(string product, string property, StringBuilder valueBuf, ref Int32 len);



    public static void GetAppsUsingAPI()
    {
        Dictionary<string, string> installedApps = new Dictionary<string, string>();

        int i = 0;
        StringBuilder productCode = new StringBuilder(39);
        Int32 productNameLen = 512;
        Int32 installLocationLen = 512;
        StringBuilder productName = new StringBuilder(productNameLen);
        StringBuilder installLocation = new StringBuilder(installLocationLen);

        while (MsiEnumProducts(i, productCode) == 0)
        {
            if (MsiGetProductInfo(productCode.ToString(), "ProductName", productName, ref productNameLen) == 0 &&
                MsiGetProductInfo(productCode.ToString(), "InstallLocation", installLocation, ref installLocationLen) == 0)
            {
                installedApps[productName.ToString()] = installLocation.ToString();
            }

            i++;
            productNameLen = 512;
            installLocationLen = 512;
            productName.Clear();
            installLocation.Clear();
        }

        foreach (var kv in installedApps)
        {
            Console.WriteLine($"Installed app: {kv.Key}, Install Location: {kv.Value}");
        }
        Console.WriteLine($"Total Unique Apps = {installedApps.Count}");

    }



    static string GetProductInfo(string productCode, string property)
    {
        StringBuilder valueBuf = new StringBuilder(1024);
        int len = 1024;

        if (MsiGetProductInfo(productCode, property, valueBuf, ref len) == 0)
        {
            Console.WriteLine(MsiGetProductInfo(productCode, property, valueBuf, ref len));

            return valueBuf.ToString();
        }
        else
        {

            return null;
        }
    }
}
