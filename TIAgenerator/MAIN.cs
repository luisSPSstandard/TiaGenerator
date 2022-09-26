using TIAgenerator.TIA_Portal;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TIAgenerator.HMI;
using TIAgenerator.PLC;

namespace TIAgenerator
{
    internal class MAIN
    {

        [STAThread]
        static void Main(string[] args)

        {

            // Result
            Boolean result = false;

            // Create new instance of class TIA_V16
            TIA_V17 newTIA = new TIA_V17();
                   
            // User dialog to open or create new project
            Console.Write("Open or create or connect to TIA project (open/create/connect): ");
            string prjCreateOpen = Console.ReadLine();

            if (prjCreateOpen == "create")
            {

                // Call method to open new TIA instance with user interface
                Console.Write("Open new TIA instance...");
                newTIA.CreateTIAinstance(true);
                Console.Write("done\n\r");


                // Get project path
                Console.Write("Project path: ");
                string prjPath = Console.ReadLine();

                // Get project name
                Console.Write("Project name: ");
                string prjName = Console.ReadLine();

                // Create new project at given path and name;
                result = newTIA.CreateTIAprj(@"" + prjPath, prjName, false);

            }
            else if (prjCreateOpen == "open")
            {

                // Call method to open new TIA instance with user interface
                Console.Write("Open new TIA instance...");
                newTIA.CreateTIAinstance(true);
                Console.Write("done\n\r");


                // Get project path
                Console.Write("Project path: ");
                string prjPath = Console.ReadLine();

                // Get project name
                Console.Write("Project name: ");
                string prjName = Console.ReadLine();

                // Create new project at given path and name;
                result = newTIA.CreateTIAprj(@"" + prjPath, prjName, true);

            } else if (prjCreateOpen == "connect")
            {
                result = newTIA.ConnectToTIA();

            }
            if (result)
            {


                HmiDevice hmi001 = newTIA.NewHmiDevice();

                Console.Write("Select device type (1 = TP1500Comfort): ");
                string devSelection = Console.ReadLine();
                Console.WriteLine("Selected device: " + devSelection);

                Console.Write("Select device name: ");
                string devName = Console.ReadLine();
                Console.WriteLine("Selected device name: " + devName);

                Console.Write("Create device...");
                hmi001.CreateDev(devName, int.Parse(devSelection));
                Console.Write("done\n\r");

                Console.WriteLine("Found software: " + hmi001.GetSoftware());

                Console.Write("Set IP address 192.168.0.123...");
                hmi001.SetIpAdress("192.168.0.123");
                Console.Write("done\n\r");

                Console.Write("Open library...");
                hmi001.OpenGlobalLibrary();
                Console.Write("done\n\r");

                Console.Write("Create projects structure...");
                hmi001.InitCreateProject();
                Console.Write("done\n\r");

                Console.Write("Close global library...");
                hmi001.CloseGlobalLibrary();
                Console.Write("done\n\r");

                Console.Write("Project creation finished!");
                Console.ReadLine();

            } else
            {

                Console.WriteLine("No valid TIA instance or no valid TIA project found. Check also, if more then one TIA instance and/or project are opened!");
                Console.ReadLine();

            }

        }
    }
}
