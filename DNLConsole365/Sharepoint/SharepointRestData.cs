using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace DNLConsole365.Sharepoint
{
   
    public class SharepointRestData
    {
        public D d { get; set; }
    }


    public class D
    {

        public string ID { get; set; }

        public string ListItemEntityTypeFullName { get; set; }
    }
}
