using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;
using System.ServiceModel.Description;

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;

namespace DNLConsole365.Service
{
    public class OrganizationService2
    {
        private static OrganizationService2 instance;

        private const string organizationServiceUri = "https://huntongroup.api.crm.dynamics.com/XRMServices/2011/Organization.svc";
        private const string username = "dynamics@huntongroup.com";
        private const string password = "CrmAdmin123!";
        
        private IOrganizationService service;

        public static OrganizationService2 Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new OrganizationService2();
                }
                return instance;
            }
        }

        private OrganizationService2()
        {
            var oUri = new Uri(organizationServiceUri);
            // Service client credentials           
            var clientCredentials = new ClientCredentials();
            clientCredentials.UserName.UserName = username;
            clientCredentials.UserName.Password = password;

            try
            {
                //Create Organization Service Proxy                  
                var serviceProxy = new OrganizationServiceProxy(oUri, null, clientCredentials, null);
                service = (IOrganizationService)serviceProxy;
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw (ex);
            }
        }

        public IOrganizationService GetService()
        {
            return service;
        }
    }
}
