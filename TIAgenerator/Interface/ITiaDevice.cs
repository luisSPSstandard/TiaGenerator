using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TIAgenerator.Interface
{
    /// <summary>
    /// Interface for device classes
    /// </summary>
    public interface ITiaDevice
    {
        /// <summary>
        /// Create new device
        /// </summary>
        /// <param name="name">device name</param>
        /// <param name="id">device identifier</param>
        void CreateDev(string name, int id);
        /// <summary>
        /// Get software name
        /// </summary>
        /// <returns>return software name as string</returns>
        string GetSoftware();

        /// <summary>
        /// Set device name
        /// </summary>
        /// <param name="name">new device name</param>
        void SetName(string name);

        /// <summary>
        /// Set device author in properties
        /// </summary>
        /// <param name="author">new author for device properties</param>
        void SetAuthor(string author);

        /// <summary>
        /// Set device comment in properties
        /// </summary>
        /// <param name="comment"></param>
        void SetComment(string comment);

        /// <summary>
        /// Set IP adress of device
        /// </summary>
        /// <param name="ipAdress">IP adress</param>
        void SetIpAdress(string ipAdress);

        /// <summary>
        /// Compile device software
        /// </summary>
        void Compile();

    }
}
