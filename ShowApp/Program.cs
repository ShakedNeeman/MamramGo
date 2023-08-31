using System;

class Program
{
    static void Main()
    {
        bool notExited = true;  // Initialize Boolean variable

        while (notExited)  // Loop will run as long as the variable is true
        {
            Console.Clear();  // Clear the console at the start of each iteration
            Console.WriteLine("Choose a method to display Windows services:");
            Console.WriteLine("1. Using Registry");
            Console.WriteLine("2. Using WMI");
            Console.WriteLine("3. Using WinAPI");
            Console.WriteLine("0. Exit");

            switch (Console.ReadLine())
            {
                case "1":
                    RegistryServiceDisplay.Display();
                    break;

                case "2":
                    WMIServiceDisplay.Display();
                    break;

                case "3":
                    WinAPIServiceDisplay.Display();
                    break;

                case "0":
                    notExited = false;  // Set variable to false to exit the loop
                    break;

                default:
                    Console.WriteLine("Invalid choice. Please try again.");
                    break;
            }

            if (notExited)
            {
                Console.WriteLine("Press any key to return to the menu...");
                Console.ReadKey();
                Console.Clear();  // Clear the console for the next iteration
            }
        }
    }
}
