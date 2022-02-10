using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DreportLauncher
{
    class DreportModels
    {

        public class DreportConfigs
        {
            public string Appversion { get; set; }
            public List<FileNames> Files { get; set; }
        }

        public class FileNames
        {
            public string file_name { get; set; }
        }
    }
}
