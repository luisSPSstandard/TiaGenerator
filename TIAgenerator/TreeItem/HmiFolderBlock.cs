using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TIAgenerator.TreeItem
{
    public class HmiFolderBlock
    {

        // Class parameters
        public string name;
        public int type; // 1 = folder; 2 = PLC block ; 3 = HMI screen
        public List<HmiFolderBlock> subhmifolderblocks;

        public HmiFolderBlock(string name, int type)
        {
            this.name = name;
            this.type = type;
            this.subhmifolderblocks = new List<HmiFolderBlock>();
        }
    }
}
