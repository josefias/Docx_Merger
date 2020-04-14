using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Doc_Merger
{
    public abstract class Paths
    {
        public static List<string> AllFilePath { get; set; }
        public List<string> FinalPath { get; set; }
        public List<string> ErrorLogPath { get; set; }
    }
}
