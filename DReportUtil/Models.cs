using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DReportUtil
{
    public class Models
    {
        public class EmailReqBody
        {
            public string toUser { get; set; }
            public string fromUser { get; set; }
            public string emailSubject { get; set; }
            public string emailBody { get; set; }
        }
    }
}
