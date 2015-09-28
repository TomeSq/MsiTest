using System;
using System.IO;
using System.Reflection;

namespace MsiTest
{
    class Program
    {
        static void Main(string[] args)
        {
            if(args.Length < 1)
            {
                string modulePath = Path.GetFileName(Assembly.GetExecutingAssembly().Location);
                Console.WriteLine(modulePath + " MSIFilePath");
                return;
            }

            MsiInstaller installer = new MsiInstaller(args[0]);

            string ProductVersion = installer.Major + "." + installer.Minor + "." + installer.Build;
            Console.WriteLine(ProductVersion);
        }
    }
}
