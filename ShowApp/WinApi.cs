using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;

class WinApi
{
    [DllImport("msi.dll", CharSet = CharSet.Unicode)]
    static extern int MsiEnumProducts(int iProductIndex, StringBuilder lpProductBuf);

    [DllImport("msi.dll", CharSet = CharSet.Unicode)]
    static extern int MsiGetProductInfo(string product, string property, StringBuilder valueBuf, ref Int32 len);

    public static void GetAppsUsingAPI()
    {
        Dictionary<string, string> appsByLocation = new Dictionary<string, string>();

        StringBuilder productCode = new StringBuilder(39);
        int index = 0;

        while (MsiEnumProducts(index, productCode) == 0)
        {
            string code = productCode.ToString();

            string productName = GetProductInfo(code, "ProductName");
            string installLocation = GetProductInfo(code, "InstallDate");

            if (!string.IsNullOrEmpty(productName))
            {
               
                    appsByLocation[installLocation] = productName;
                
            }

            index++;
        }

        foreach (var pair in appsByLocation)
        {
            Console.WriteLine($"Product: {pair.Key}, Location: {pair.Value}");
        }
        Console.WriteLine($"Total Unique Apps = {appsByLocation.Count}");

    }

    static string GetProductInfo(string productCode, string property)
    {
        StringBuilder valueBuf = new StringBuilder(512);
        int len = 512;

        if (MsiGetProductInfo(productCode, property, valueBuf, ref len) == 0)
        {
            return valueBuf.ToString();
        }
        else
        {
            return null;
        }
    }
}
