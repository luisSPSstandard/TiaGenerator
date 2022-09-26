using Newtonsoft.Json;
using Siemens.Engineering;
using Siemens.Engineering.Compiler;
using Siemens.Engineering.Hmi;
using Siemens.Engineering.Hmi.Screen;
using Siemens.Engineering.HW;
using Siemens.Engineering.HW.Features;
using Siemens.Engineering.Library;
using Siemens.Engineering.Library.MasterCopies;
using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TIAgenerator.Interface;
using TIAgenerator.TreeItem;

namespace TIAgenerator.HMI
{
    /// <summary>
    /// HMI device class implementing ITiaDevice interface
    /// </summary>
    public class HmiDevice : ITiaDevice
    {

        private static TiaPortal instTIA;
        private static Project projectTIA;
        private static Device hmiDevice;
        private static DeviceItem hmiDeviceItem;
        private static HmiTarget hmiSwTarget;
        private static UserGlobalLibrary glbLib;
        private static NetworkInterface hmiNetInterface;
        private static Node hmiNode;
        private TiaPortal instTIA1;
        private Project projectTIA1;

        /// <summary>
        /// HMI device class constructor
        /// </summary>
        /// <param name="instance">Running TIA Portal instance</param>
        /// <param name="project">Opened TIA Portal project</param>
        public HmiDevice(TiaPortal instance, Project project)
        {
            instTIA = instance;
            projectTIA = project;
        }

        /// <summary>
        /// Create new hardware device
        /// </summary>
        /// <param name="devName">Device name</param>
        /// <param name="devSelection">Device type</param>
        public void CreateDev(string devName, int devSelection)
        {

            string devVersion = "";
            string devArticle = "";
            string devStation = "station" + devName;


            switch (devSelection)
            {

                case 1: // TP1500 Comfort
                    devArticle = "6AV2 124-0QC02-0AX1";
                    devVersion = "16.0.0.0";
                    devStation = null;
                    break;

            }

            string devIdent = "OrderNumber:" + devArticle + "/" + devVersion;

            // Create new device within project
            hmiDevice = projectTIA.Devices.CreateWithItem(devIdent, devName, devStation);

            // Get PLC device item
            hmiDeviceItem = hmiDevice.DeviceItems.First(Device => Device.Name.Equals(devName));

        }

        /// <summary>
        /// Get HMI software by HMI target
        /// </summary>
        /// <returns>Software as string</returns>
        public string GetSoftware()
        {
            foreach (DeviceItem item in hmiDevice.DeviceItems)
            {

                SoftwareContainer hmiSwContainer = ((IEngineeringServiceProvider)item).GetService<SoftwareContainer>();

                if (hmiSwContainer != null)
                {

                    Software softwareHmi = hmiSwContainer.Software;
                    hmiSwTarget = softwareHmi as Siemens.Engineering.Hmi.HmiTarget;
                    return hmiSwTarget.Name;
                }
            }

            return null;
        }

        /// <summary>
        /// Set IP adress of HMI device
        /// </summary>
        /// <param name="ipAdress">IP adress for HMI device</param>
        public void SetIpAdress(string ipAdress)
        {

            hmiNetInterface = null;

            //Serach for CP devices
            foreach (DeviceItem item in hmiDevice.DeviceItems)
            {

                // Search for interace
                foreach (DeviceItem hmiInterface in item.DeviceItems)
                {

                    if (hmiInterface.Name == "PROFINET Schnittstelle_1")
                    {

                        // Set HMI interface object
                        hmiNetInterface = ((IEngineeringServiceProvider)hmiInterface).GetService<NetworkInterface>();
                        break;

                    }

                }

                // Break if HMI interface found
                if (hmiNetInterface != null)
                {
                    break;
                }
            }

            if (hmiNetInterface != null)
            {
                // Search for proper node
                foreach (Node node in hmiNetInterface.Nodes)
                {

                    if (node != null)
                    {

                        // Check if node info with name "Address" exists
                        foreach (EngineeringAttributeInfo nodeInfo in node.GetAttributeInfos())
                        {

                            if (nodeInfo.Name == "Address" && ipAdress != null)
                            {

                                // Set IP address of current node
                                node.SetAttribute("Address", ipAdress);
                                hmiNode = node;
                            }

                        }

                    }

                }

            }

        }
        /// <summary>
        /// Compile HMI device
        /// </summary>
        public void Compile()
        {
            // Get compilable object from plc
            ICompilable plcCompilable = hmiDevice.GetService<ICompilable>();

            // Exceute compile
            CompilerResult plcCompileResult = plcCompilable.Compile();

            // Print compile results
            PrintCompileResult(plcCompileResult);
        }

        /// <summary>
        /// Call printing methode for compile results
        /// </summary>
        /// <param name="result">Compiler result</param>
        private void PrintCompileResult(CompilerResult result)
        {

            Console.WriteLine("State: " + result.State);
            Console.WriteLine("Errors: " + result.ErrorCount);
            Console.WriteLine("Warnings: " + result.WarningCount);
            RecursiveCompileResult(result.Messages);

        }
        /// <summary>
        /// Recursive call to print compile result
        /// </summary>
        /// <param name="messages">Compiler result</param>
        private void RecursiveCompileResult(CompilerResultMessageComposition messages)
        {

            foreach (CompilerResultMessage result in messages)
            {

                // Print result details
                Console.WriteLine("Path: " + result.Path);
                Console.WriteLine("Date and Time: " + result.DateTime);
                Console.WriteLine("Description: " + result.Description);

                // Call further results
                RecursiveCompileResult(result.Messages);
            }


        }

        /// <summary>
        /// Open global library
        /// </summary>
        public void OpenGlobalLibrary()
        {
            // Create new OpenFileDialog object
            using (OpenFileDialog fileDialog = new OpenFileDialog())
            {

                // Set search filter
                fileDialog.InitialDirectory = "C:\\Users\\Admin\\Documents\\Automatisierung";
                fileDialog.Filter = "TIA library files (*.al17)|*.al17";

                if (fileDialog.ShowDialog() == DialogResult.OK)
                {
                    // Get global library path
                    System.IO.FileInfo libPath = new System.IO.FileInfo(@"" + fileDialog.FileName);

                    // Open global library
                    glbLib = instTIA.GlobalLibraries.Open(libPath, OpenMode.ReadOnly);
                }

            }

        }


        /// <summary>
        /// Close global library
        /// </summary>        
        public void CloseGlobalLibrary()
        {

            // Close global library
            glbLib.Close();

        }
        
        /// <summary>
        /// Init methode to create a project based on JSON
        /// </summary>
        public void InitCreateProject()
        {

            using (OpenFileDialog fileDialog = new OpenFileDialog())
            {

                HmiFolderBlock treeRoot;

                // Set search filter
                fileDialog.InitialDirectory = "C:\\Users\\Admin\\Documents\\Automatisierung";
                fileDialog.Filter = "JSON|*.json";

                if (fileDialog.ShowDialog() == DialogResult.OK)
                {

                    // Create stream reader to read json file
                    using (StreamReader readFile = new StreamReader(fileDialog.FileName))
                    {

                        string json = readFile.ReadToEnd();

                        // Convert json to Folder
                        treeRoot = JsonConvert.DeserializeObject<HmiFolderBlock>(json);
                        CreateProjectStruct(treeRoot, null);
                    }

                }

            }

        }
        
        /// <summary>
        /// Create project structure
        /// </summary>
        /// <param name="treeRoot">Tree node deserialized from JSON </param>
        /// <param name="hmiFolder">Current folder within HMI screens</param>
        public void CreateProjectStruct(HmiFolderBlock treeRoot, ScreenFolder hmiFolder)
        {
            // Create local variable for logical folder
            ScreenFolder currFolder = hmiFolder;

            if (treeRoot.type == 1)
            {
                // Create subfolder
                if (hmiFolder != null)
                {

                    currFolder = hmiFolder.Folders.Create(treeRoot.name);

                }

                // Create main folder
                else
                {

                    currFolder = hmiSwTarget.ScreenFolder.Folders.Create(treeRoot.name);

                }
            }

            if (treeRoot.type == 3)
            {
                // Create HMI screen at given folder from mastercopy
                GetMasterCopy(glbLib.MasterCopyFolder, hmiFolder, treeRoot.name);

            }


            if (treeRoot.subhmifolderblocks.Count > 0)
            {
                // Resursive method call for each subnode of current tree node
                foreach (HmiFolderBlock subfolderblock in treeRoot.subhmifolderblocks)
                {

                    Console.WriteLine("Call creating method for current subfolder...");
                    CreateProjectStruct(subfolderblock, currFolder);

                }
            }
        }


        /// <summary>
        /// Copy an HMI screen from global library at given folder path. The HMI screen is selected by name read
        /// </summary>
        /// <param name="folder">Current copy folder from global lib</param>
        /// <param name="hmiFolder">Current HMI folder</param>
        /// <param name="hmiName">HMI name to be copied from library</param>
        public void GetMasterCopy(MasterCopyFolder folder, ScreenFolder hmiFolder, string hmiName)
        {

            // Call recursive function for subfolders
            if (folder.Folders != null)
            {

                foreach (MasterCopyFolder copyFolder in folder.Folders)
                {
                    Console.WriteLine("Call GetMasterCopy for subfolder " + copyFolder.Name);
                    GetMasterCopy(copyFolder, hmiFolder, hmiName);

                }

            }

            // Create master copy if availabe
            if (folder.MasterCopies != null)
            {
                foreach (MasterCopy masterCopy in folder.MasterCopies)
                {

                    // Create HMI screen from mastercopy if names match and logical folder valid
                    if (hmiName == masterCopy.Name && hmiFolder != null)
                    {
                        Console.WriteLine("Create " + masterCopy.Name + " at " + hmiFolder.Name + "...");
                        hmiFolder.Screens.CreateFrom(masterCopy);

                        // Create HMI screen from mastercopy if names match an no logical folder given
                    }
                    else if (hmiName == masterCopy.Name && hmiFolder == null)
                    {

                        Console.WriteLine("Create " + masterCopy.Name + " at highest level...");
                        hmiSwTarget.ScreenFolder.Screens.CreateFrom(masterCopy);

                    }
                }
            }
        }


        /// <summary>
        /// Set author im HMI device properties
        /// </summary>
        /// <param name="author">Author name</param>
        public void SetAuthor(string author)
        {
            hmiDevice.SetAttribute("Author", author);
        }

        /// <summary>
        /// Set comment in HMI device properties
        /// </summary>
        /// <param name="comment">Comment string</param>
        public void SetComment(string comment)
        {
            hmiDevice.SetAttribute("Comment", comment);
        }

        /// <summary>
        /// Set name in HMI device properties
        /// </summary>
        /// <param name="name">Device name</param>
        public void SetName(string name)
        {
            hmiDevice.SetAttribute("Name", name);
        }
    }

}
