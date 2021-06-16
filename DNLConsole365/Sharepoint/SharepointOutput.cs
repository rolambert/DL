using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DNLConsole365.Sharepoint
{
    public class SharePointOutput
    {
        public bool Success { get; set; } = true;

        public Exception Exception { get; set; } = null;

        public string ListItemId { get {
                return SPRestData == null ? "-1" : SPRestData.d.ID;
            } }

        public SharepointRestData SPRestData { get; set; }
        public string JsonRestData { get; set; }

    }
}
