using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System.Net;
using System.IO;
using System.Security.Cryptography.X509Certificates;


namespace DNLConsole365.Projects
{
    public class RightMoveTest
    {

        public static List<Entity> GetWebResources(IOrganizationService service)
        {
            var query = new QueryExpression("contact");
            query.ColumnSet = new ColumnSet(true);

            return service.RetrieveMultiple(query).Entities.ToList();
        }
        public static Entity GetStoredCertificate(IOrganizationService service)
        {
            var query = new QueryExpression()
            {
                EntityName = "webresource",
                ColumnSet = new ColumnSet("content"),
                Criteria = new FilterExpression
                {
                    FilterOperator = LogicalOperator.And,
                    Conditions =
                     {
                        new ConditionExpression
                        {
                            AttributeName = "name",
                            Operator = ConditionOperator.Equal,
                            Values = { "new_rightmovecert" }
                        }
                    }
                }
            };

            return service.RetrieveMultiple(query).Entities.FirstOrDefault();
        }

        
        public static string MakeRequest(string data, string method, byte[] certBytes)
        {
            var cert3 = new X509Certificate2(certBytes, "FHiYXCFNsv");

            var httpWebRequest = (HttpWebRequest)WebRequest.Create("https://adfapi.adftest.rightmove.com/v1/property/" + method);
            httpWebRequest.ContentType = "application/json";
            httpWebRequest.Method = "POST";
            httpWebRequest.ClientCertificates.Clear();
            httpWebRequest.ClientCertificates.Add(cert3);
            using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
            {
                streamWriter.Write(data);
                streamWriter.Flush();
                streamWriter.Close();
            }

            var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
            using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
            {
                var result = streamReader.ReadToEnd();
                return result;
            }
        }    
    }
}
