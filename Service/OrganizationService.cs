using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;
using System.ServiceModel.Description;
using System.Net;

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;

namespace DNLConsole365.Service
{
    public class OrganizationCredentials
    {
        public string OrganizationUrl { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }

        public OrganizationCredentials(string uri, string user, string pass)
        {
            OrganizationUrl = uri;
            UserName = user;
            Password = pass;
        }
    }

    public class OrganizationService
    {
        private static OrganizationService instance;

        /*private const string organizationServiceUri = "https://legal.api.crm.dynamics.com/XRMServices/2011/Organization.svc";
        private const string username = "vlad@forceworks.com";
        private const string password = "Forceworks16";*/

        /*private const string organizationServiceUri = "https://crm.huntongroup.com/XRMServices/2011/Organization.svc";
        private const string username = "HUNTON\\cteam";
        private const string password = "Trane1";*/        

        private IOrganizationService service;

        public static OrganizationService Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new OrganizationService();
                }
                return instance;
            }
        }

        private OrganizationService()
        {
         
        }

        public IOrganizationService GetService(OrganizationCredentials creds, Guid _callerId)
        {
            var oUri = new Uri(creds.OrganizationUrl + "/XRMServices/2011/Organization.svc");
            // Service client credentials           
            var clientCredentials = new ClientCredentials();
            clientCredentials.UserName.UserName = creds.UserName;
            clientCredentials.UserName.Password = creds.Password;            

            try
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                //Create Organization Service Proxy                  
                using (var serviceProxy = new OrganizationServiceProxy(oUri, null, clientCredentials, null))
                {                    
                    if (_callerId != Guid.Empty)
                    {
                        serviceProxy.CallerId = _callerId;
                    }
                    serviceProxy.Timeout = new TimeSpan(0, 15, 0);

                    return (IOrganizationService)serviceProxy;

                }
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw (ex);
            }            
        }
    }
}
