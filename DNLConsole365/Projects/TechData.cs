using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.ServiceModel.Description;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Tooling.Connector;

namespace DNLConsole365.Projects
{
    public class TechData
    {
        public static CrmServiceClient CreateCrmConnection(string userName, string password)
        {
            var url = "Url=https://techdata.api.crm.dynamics.com;AuthType=Office365;";// ConfigurationManager.ConnectionStrings["CrmService"].ConnectionString;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            var client = new CrmServiceClient(string.Format("{0}UserName={1};Password={2};", url, userName, password));

            if (client.IsReady)
            {
                return client;
            }
            else
            {
                // Display the last error.
                Console.WriteLine("Error occurred: {0}", client.LastCrmError);

                // Display the last exception message if any.
                Console.WriteLine(client.LastCrmException.Message);
                Console.WriteLine(client.LastCrmException.Source);
                Console.WriteLine(client.LastCrmException.StackTrace);

                throw new Exception("Unable to Connect to CRM");
            }
        }

        internal static OrganizationServiceProxy GetService(string login, string password, string url)
        {

            var organizationUri = new Uri("https://tdusdev.api.crm.dynamics.com/XRMServices/2011/Organization.svc");
            var credentials = new ClientCredentials();
            credentials.UserName.UserName = login;
            credentials.UserName.Password = password;
            credentials.Windows.ClientCredential = new NetworkCredential(login, password);
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            return new OrganizationServiceProxy(organizationUri, null, credentials, null);
        }
    }
}
