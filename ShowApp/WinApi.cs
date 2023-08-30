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
            Dictionary<string, string> appsByVersion = new Dictionary<string, string>();

            StringBuilder productCode = new StringBuilder(39);
            int index = 0;

            while (MsiEnumProducts(index, productCode) == 0)
            {
                string code = productCode.ToString();

                string productName = GetProductInfo(code, "ProductName");
                string version = GetProductInfo(code, "VersionString");

                if (!string.IsNullOrEmpty(productName) && !string.IsNullOrEmpty(version))
                {
                    appsByVersion[version] = productName;
                }

                index++;
            }

            Console.WriteLine($"Total Apps = {index}");

            foreach (var pair in appsByVersion)
            {
                Console.WriteLine($"Version: {pair.Key}, Product: {pair.Value}");
            }
            Console.WriteLine($"Total Unique Apps = {appsByVersion.Count}");

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

