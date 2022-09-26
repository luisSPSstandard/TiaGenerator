using Microsoft.Win32;
using Newtonsoft.Json;
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
using TIAgenerator.HMI;
using TIAgenerator.PLC;

namespace TIAgenerator.TIA_Portal
{
    /// <summary>
    /// TIA class
    /// </summary>
    public class TIA_V17
    {

        private static TiaPortal instTIA;
        private static Project projectTIA;

        /// <summary>
        /// Open a new instance of TIA Portal V17 with/without user interface
        /// </summary>
        /// <param name="guiTIA">Open TIA Portal with GUI</param>
        public void CreateTIAinstance(bool guiTIA)
        {

            // Set whitelist entry
            SetWhitelist(System.Diagnostics.Process.GetCurrentProcess().ProcessName, System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName);

            // Open new TIA instance with user interface
            if (guiTIA)
            {
                instTIA = new TiaPortal(TiaPortalMode.WithUserInterface);

            }

            // Open TIA instance without user interface
            else
            {

                instTIA = new TiaPortal(TiaPortalMode.WithoutUserInterface);

            }

        }

        /// <summary>
        /// Create or open TIA project
        /// </summary>
        /// <param name="prjPath">TIA Portal project path</param>
        /// <param name="prjName">TIA Portal project name</param>
        /// <param name="open">Open existing project at given path with given name. Else create new project</param>
        /// <returns></returns>
        public Boolean CreateTIAprj(string prjPath, string prjName, bool open)
        {

            if (!open)
            {

                // Create new directory info
                DirectoryInfo targetDir = new DirectoryInfo(prjPath);

                // Create new TIA project
                projectTIA = instTIA.Projects.Create(targetDir, prjName);


            }
            else
            {

                // Create new file info
                FileInfo targetDir = new FileInfo(prjPath + "\\" + prjName + ".ap16");

                // Open exisitng TIA project
                projectTIA = instTIA.Projects.Open(targetDir);

            }

            // Check if project is valid
            if (projectTIA != null)
            {

                return true;
            }

            return false;
        }

        /// <summary>
        /// Connect to a running TIA Portal instance. Only one TIA Portal may be open at the same time. 
        /// </summary>
        /// <returns>True, if connected</returns>
        public Boolean ConnectToTIA()
        {

            // Set whitelist entry
            SetWhitelist(System.Diagnostics.Process.GetCurrentProcess().ProcessName, System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName);

            // Get all TIA processes
            IList<TiaPortalProcess> tiaProcesses = TiaPortal.GetProcesses();

            // Check if only one instance of TIA is open and connect to this instance
            switch (tiaProcesses.Count)
            {

                case 1:
                    instTIA = tiaProcesses[0].Attach();
                    break;

                default:
                    instTIA = null;
                    break;

            }

            // Get the currently opened project if there is a valid TIA instance
            if (instTIA != null)
            {

                projectTIA = instTIA.Projects[0];

                if (projectTIA != null)
                {

                    return true;

                }

            }

            return false;

        }

        /// <summary>
        /// Create new element from PlcDevice class
        /// </summary>
        /// <returns></returns>
        public PlcDevice NewPlcDevice()
        {
            PlcDevice device = new PlcDevice(instTIA, projectTIA);
            return device;
        }

        /// <summary>
        /// Create new element from HmiDevice class
        /// </summary>
        /// <returns></returns>
        public HmiDevice NewHmiDevice()
        {
            HmiDevice device = new HmiDevice(instTIA, projectTIA);
            return device;
        }


        /// <summary>
        /// Set whitelist entry for TIA Portal in registry
        /// </summary>
        /// <param name="ApplicationName"></param>
        /// <param name="ApplicationStartupPath"></param>
        public void SetWhitelist(string ApplicationName, string ApplicationStartupPath)
        {

            RegistryKey key = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
            RegistryKey software = null;
            try
            {
                software = key.OpenSubKey(@"SOFTWARE\Siemens\Automation\Openness")
                    .OpenSubKey("17.0")
                    .OpenSubKey("Whitelist")
                    .OpenSubKey(ApplicationName + ".exe")
                    .OpenSubKey("Entry", RegistryKeyPermissionCheck.ReadWriteSubTree, RegistryRights.FullControl);
            }
            catch (Exception)
            {

                //Eintrag in der Whitelist ist nicht vorhanden
                //Entry in whitelist is not available
                software = key.CreateSubKey(@"SOFTWARE\Siemens\Automation\Openness")
                    .CreateSubKey("17.0")
                    .CreateSubKey("Whitelist")
                    .CreateSubKey(ApplicationName + ".exe")
                    .CreateSubKey("Entry", RegistryKeyPermissionCheck.ReadWriteSubTree, RegistryOptions.None);
            }


            string lastWriteTimeUtcFormatted = String.Empty;
            DateTime lastWriteTimeUtc;
            HashAlgorithm hashAlgorithm = SHA256.Create();
            FileStream stream = File.OpenRead(ApplicationStartupPath);
            byte[] hash = hashAlgorithm.ComputeHash(stream);
            // this is how the hash should appear in the .reg file
            string convertedHash = Convert.ToBase64String(hash);
            software.SetValue("FileHash", convertedHash);
            lastWriteTimeUtc = new FileInfo(ApplicationStartupPath).LastWriteTimeUtc;
            // this is how the last write time should be formatted
            lastWriteTimeUtcFormatted = lastWriteTimeUtc.ToString(@"yyyy/MM/dd HH:mm:ss.fff", CultureInfo.InvariantCulture);
            software.SetValue("DateModified", lastWriteTimeUtcFormatted);
            software.SetValue("Path", ApplicationStartupPath);

        }

    }

}
