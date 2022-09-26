using Microsoft.Win32;
using Newtonsoft.Json;
using TIAgenerator.TreeItem;
using Siemens.Engineering;
using Siemens.Engineering.Compiler;
using Siemens.Engineering.Connection;
using Siemens.Engineering.Download;
using Siemens.Engineering.Download.Configurations;
using Siemens.Engineering.Hmi;
using Siemens.Engineering.Hmi.Screen;
using Siemens.Engineering.HW;
using Siemens.Engineering.HW.Features;
using Siemens.Engineering.Library;
using Siemens.Engineering.Library.MasterCopies;
using Siemens.Engineering.SW;
using Siemens.Engineering.SW.Blocks;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.AccessControl;
using System.Security.Cryptography;
using System.Windows.Forms;
using TIAgenerator.Interface;

namespace TIAgenerator.PLC
{   
    /// <summary>
    /// PLC device class implementing ITiaDevice interface
    /// </summary>
    public class PlcDevice : ITiaDevice
    {

        private static TiaPortal instTIA;
        private static Project projectTIA;
        private static Device plcDevice;
        private static DeviceItem plcDeviceItem;
        private static PlcSoftware softwarePLC;
        private static UserGlobalLibrary glbLib;
        private static NetworkInterface plcNetInterface;
        private static Node plcNode;

        /// <summary>
        /// Constructor for PlcDevice class
        /// </summary>
        /// <param name="instance">Current TIA Portal instance</param>
        /// <param name="project">Current TIA Portal project</param>
        public PlcDevice(TiaPortal instance, Project project)
        {

            instTIA = instance;
            projectTIA = project;

        }

        /// <summary>
        /// Create new PLC device
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

                case 1: //S7-314C-2
                    devArticle = "6ES7 314-6EH04-0AB0";
                    devVersion = "V3.3";
                    break;

                case 2: // S7-1518F-4
                    devArticle = "6ES7 518-4AP00-0AB0";
                    devVersion = "V2.8";
                    break;


            }

            string devIdent = "OrderNumber:" + devArticle + "/" + devVersion;

            // Create new device within project
            plcDevice = projectTIA.Devices.CreateWithItem(devIdent, devName, devStation);

            // Get PLC device item
            plcDeviceItem = plcDevice.DeviceItems.First(Device => Device.Name.Equals(devName));

        }

        /// <summary>
        /// Get PLC software
        /// </summary>
        /// <returns>Software as string</returns>
        public string GetSoftware()
        {
            // Iterate through device items to search for software element
            foreach (DeviceItem devItem in plcDevice.DeviceItems)
            {

                SoftwareContainer plcSWcontainer = ((IEngineeringServiceProvider)devItem).GetService<SoftwareContainer>();

                // Parse found software to the type PLC software
                if (plcSWcontainer != null)
                {

                    softwarePLC = (PlcSoftware)plcSWcontainer.Software;
                    return softwarePLC.ToString();

                }
            }

            return null;

        }

        /// <summary>
        /// Set name in PLC device properties
        /// </summary>
        /// <param name="plcName">Device name</param>
        public void SetName(string plcName)
        {
            // Iterate through device items to search for software element
            foreach (DeviceItem devItem in plcDevice.DeviceItems)
            {

                if (devItem != null)
                {

                    // Set PLC name 
                    devItem.SetAttribute("Name", plcName);

                }

            }
        }

        /// <summary>
        /// Set Author in PLC device properties
        /// </summary>
        /// <param name="plcAuthor"></param>
        public void SetAuthor(string plcAuthor)
        {
            // Iterate through device items to search for software element
            foreach (DeviceItem devItem in plcDevice.DeviceItems)
            {

                if (devItem != null)
                {

                    // Set PLC author
                    devItem.SetAttribute("Author", plcAuthor);

                }

            }
        }

        /// <summary>
        /// Set author in PLC device properties
        /// </summary>
        /// <param name="plcName">PLC comment</param>
        /// <returns></returns>
        public void SetComment(string plcComment)
        {
            // Iterate through device items to search for software element
            foreach (DeviceItem devItem in plcDevice.DeviceItems)
            {

                if (devItem != null)
                {

                    // Set PLC author
                    devItem.SetAttribute("Comment", plcComment);

                }
            }
        }
        
        /// <summary>
        /// Set IP adress in PLC device properties
        /// </summary>
        /// <param name="ipAdress">IP adress</param>
        public void SetIpAdress(string ipAdress)
        {

            // Get first PROFINET interface
            DeviceItem plcProfinet = plcDeviceItem.DeviceItems.First(DeviceItem => DeviceItem.Name.Equals("PROFINET-Schnittstelle_1"));

            // Get network interface for plc node
            plcNetInterface = ((IEngineeringServiceProvider)plcProfinet).GetService<NetworkInterface>();

            if (plcNetInterface != null)
            {
                // Get network node
                foreach (Node node in plcNetInterface.Nodes)
                {

                    Console.WriteLine("Found node: " + node.Name);

                    if (node != null)
                    {
                        // Search through nodes for node with IP-Adress attribute
                        foreach (EngineeringAttributeInfo nodeInfo in node.GetAttributeInfos())
                        {

                            //Console.WriteLine("Node attributes: " + nodeInfo.Name);

                            if (nodeInfo != null && nodeInfo.Name == "Address" && ipAdress != null)
                            {
                                // Set IP Adress
                                node.SetAttribute("Address", ipAdress);
                                plcNode = node;

                            }

                        }

                    }

                }

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
        /// Copy PLC block from global library into PLC programming blocks. The PLC block is selected by the given name
        /// </summary>
        /// <param name="folder">Current master copy folder from global library</param>
        /// <param name="tiaFolder">Current PLC user folder in </param>
        /// <param name="blockName">PLC block name to be copied from global library</param>        
        public void GetMasterCopy(MasterCopyFolder folder, PlcBlockUserGroup tiaFolder, string blockName)
        {

            // Call recursive function for subfolders
            if (folder.Folders != null)
            {

                foreach (MasterCopyFolder copyFolder in folder.Folders)
                {
                    Console.WriteLine("Call GetMasterCopy for subfolder " + copyFolder.Name);
                    GetMasterCopy(copyFolder, tiaFolder, blockName);

                }

            }

            // Create master copy if availabe
            if (folder.MasterCopies != null)
            {
                foreach (MasterCopy masterCopy in folder.MasterCopies)
                {

                    // Create block from mastercopy if names match and logical folder valid
                    if (blockName == masterCopy.Name && tiaFolder != null)
                    {
                        Console.WriteLine("Create " + masterCopy.Name + " at " + tiaFolder.Name + "...");
                        tiaFolder.Blocks.CreateFrom(masterCopy);

                        // Create block from mastercopy if names match an no logical folder given
                    }
                    else if (blockName == masterCopy.Name && tiaFolder == null)
                    {

                        Console.WriteLine("Create " + masterCopy.Name + " at highest level...");
                        softwarePLC.BlockGroup.Blocks.CreateFrom(masterCopy);

                    }
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
        /// Init methode to create a project struct based in JSON
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
        /// Recursive methode to create folder tree in TIA Portal
        /// </summary>
        /// <param name="treeRoot"></param>
        /// <param name="tiafolder"></param>
        public void CreateProjectStruct(HmiFolderBlock treeRoot, PlcBlockUserGroup tiafolder)
        {
            // Create local variable for logical folder
            PlcBlockUserGroup currFolder = tiafolder;

            if (treeRoot.type == 1)
            {
                // Create subfolder
                if (tiafolder != null)
                {

                    currFolder = tiafolder.Groups.Create(treeRoot.name);

                }

                // Create main folder
                else
                {

                    currFolder = softwarePLC.BlockGroup.Groups.Create(treeRoot.name);

                }
            }

            if (treeRoot.type == 2)
            {
                // Create PLC block at given folder from mastercopy
                GetMasterCopy(glbLib.MasterCopyFolder, tiafolder, treeRoot.name);

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
        /// Create new subnet at created PLC
        /// </summary>
        public void CreateNetwork(String subnetname)
        {

            foreach (IoController controller in plcNetInterface.IoControllers)
            {

                Console.WriteLine("Found IOcontroller: " + controller.ToString());

            }

            //Get connected subnets
            Subnet plcSubnet = plcNode.ConnectedSubnet;

            if (plcSubnet == null)
            {

                plcSubnet = plcNode.CreateAndConnectToSubnet(subnetname);

            }

        }

        /// <summary>
        /// Method to compile PLC
        /// </summary>
        public void Compile()
        {

            // Get compilable object from plc
            ICompilable plcCompilable = plcDevice.GetService<ICompilable>();

            // Exceute compile
            CompilerResult plcCompileResult = plcCompilable.Compile();

            // Print compile results
            PrintCompileResult(plcCompileResult);


        }
        /// <summary>
        /// Call printing methode for compile results
        /// </summary>
        /// <param name="result"></param>
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
        /// <param name="messages"></param>
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

        public void Load()
        {

            // Get download provider
            DownloadProvider downloadProvider = plcDeviceItem.GetService<DownloadProvider>();

            // Get connection configuration
            ConnectionConfiguration plcConnectConfig = downloadProvider.Configuration;

            // Set download provider configuration
            ConfigurationMode plcConfigMode = plcConnectConfig.Modes.Find("PN/IE");
            ConfigurationPcInterface pcInterface = plcConfigMode.PcInterfaces.Find("PLCSIM", 1);
            IConfiguration plcTargetConfig = pcInterface.TargetInterfaces.Find("1 X1");

            // Load PLC
            downloadProvider.Download(plcTargetConfig, PreLoadConfig, PostLoadConfig, DownloadOptions.Hardware | DownloadOptions.Software);

        }

        private static void PreLoadConfig(DownloadConfiguration loadConfig)
        {

            StopModules stopCPU = loadConfig as StopModules;
            if (stopCPU != null)
            {

                stopCPU.CurrentSelection = StopModulesSelections.StopAll;

            }

            AllBlocksDownload loadAllBlock = loadConfig as AllBlocksDownload;
            if (loadAllBlock != null)
            {

                loadAllBlock.CurrentSelection = AllBlocksDownloadSelections.DownloadAllBlocks;

            }

            OverwriteSystemData overwriteSystemData = loadConfig as OverwriteSystemData;
            if (overwriteSystemData != null)
            {

                overwriteSystemData.CurrentSelection = OverwriteSystemDataSelections.Overwrite;

            }

            ConsistentBlocksDownload consistendLoad = loadConfig as ConsistentBlocksDownload;
            if (consistendLoad != null)
            {

                consistendLoad.CurrentSelection = ConsistentBlocksDownloadSelections.ConsistentDownload;


            }

        }

        private static void PostLoadConfig(DownloadConfiguration loadConfig)
        {

            StartModules startCPU = loadConfig as StartModules;
            if (startCPU != null)
            {

                startCPU.CurrentSelection = StartModulesSelections.StartModule;

            }

        }


    }

}
