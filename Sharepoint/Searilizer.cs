using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.IO;
using System.Runtime.Serialization.Json;

namespace DNLConsole365.Sharepoint
{
    public static class Searilizer
    {
        public static SharepointRestData DeserilizeSPJsonString(string jsonString)
        {
            SharepointRestData obj;
            try
            {
                using (var ms = new MemoryStream(Encoding.Unicode.GetBytes(jsonString)))
                {
                    // Deserialization from JSON  
                    DataContractJsonSerializer deserializer = new DataContractJsonSerializer(typeof(SharepointRestData));
                    obj = (SharepointRestData)deserializer.ReadObject(ms);
                    return obj;
                }
            }
            catch
            {
                throw;
            }
            
        }
    }
}
