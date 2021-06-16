using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DReport
{
    class allListMdl
    {
        public virtual string cardNo { get; set; }
        public virtual string beginBalance { get; set; }
        public virtual string endBalance { get; set; }
        public virtual string amount { get; set; }
        public virtual string createDate { get; set; }
        public virtual string createUser { get; set; }
        public virtual string regDate { get; set; }
    }
}
