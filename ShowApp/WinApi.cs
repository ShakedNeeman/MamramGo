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


    public static Dictionary<string, Tuple<string, string, string, string>> GetAppsUsingAPI()
    {
        // Initialize a dictionary to hold application details by package code
        Dictionary<string, Tuple<string, string, string, string>> appsByAppName = new Dictionary<string, Tuple<string, string, string, string>>();

        // StringBuilder to store the product code
        StringBuilder productCode = new StringBuilder(39);

        // Initialize index for enumeration
        int index = 0;

        // Enumerate all installed products
        while (MsiEnumProducts(index, productCode) == 0)
        {
            // Convert the product code to a string
            string code = productCode.ToString();

            // Fetch product name for the given product code
            string productName = GetProductInfo(code, "ProductName");

            // Fetch package code for the given product code
            string packageCode = GetProductInfo(code, "PackageCode") as string;

            // Fetch install date for the given product code
            string installDate = GetProductInfo(code, "InstallDate") as string;

            // Fetch installed size for the given product code
            string installedSize = GetProductInfo(code, "InstalledSize");

            // Check if the product name and package code are not null or empty
            if (!string.IsNullOrEmpty(productName) && !string.IsNullOrEmpty(packageCode))
            {
                // Create a Tuple for additional info like install date and installed size
                var additionalInfo = new Tuple<string, string, string, string>(code, packageCode, installDate ?? "N/A", installedSize ?? "N/A");

                // Add the application info to the dictionary
                appsByAppName[productName] = additionalInfo;
            }

            // Increment the index for the next iteration
            index++;
        }

        // Return the dictionary containing all the apps and their details
        return appsByAppName;
    }




    static string GetProductInfo(string productCode, string property)
    {
        StringBuilder valueBuf = new StringBuilder(1024);
        int len = 1024;

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
