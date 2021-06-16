using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MimeKit.Utils;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System.Net;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Xrm.Sdk.WebServiceClient;
using Microsoft.Crm.Sdk.Messages;
using System.Net.Http;

namespace DNLConsole365.Projects
{
    public class NorthJohnson
    {

        /// Method-to-generate-Access-Token  
        public static async Task<string> AccessTokenGenerator()
        {
            string clientId = "f3769397-a593-488f-8233-2048f9ae139e";
            string clientSecret = "__ygP9--d2C.0gX_QhyPcrC22l1XXc-5pf";
            string authority = "https://login.microsoftonline.com/46806687-f2d4-4073-bee0-a5e00bc79c27";
            string resourceUrl = "https://nj2020sandbox.crm.dynamics.com"; // Org URL  

            ClientCredential credentials = new ClientCredential(clientId, clientSecret);
            var authContext = new Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext(authority);
            var result = await authContext.AcquireTokenAsync(resourceUrl, credentials);
            return result.AccessToken;
        }

        public static async Task<HttpResponseMessage> CrmRequest(HttpMethod httpMethod, string requestUri, string body = null)
        {
            var accessToken = await AccessTokenGenerator();
            var client = new HttpClient();
            var msg = new HttpRequestMessage(httpMethod, requestUri);
            msg.Headers.Add("OData-MaxVersion", "4.0");
            msg.Headers.Add("OData-Version", "4.0");
            msg.Headers.Add("Prefer", "odata.include-annotations=\"*\"");

            // Passing AccessToken in Authentication header  
            msg.Headers.Add("Authorization", $"Bearer {accessToken}");

            if (body != null)
                msg.Content = new StringContent(body, UnicodeEncoding.UTF8, "application/json");

            return await client.SendAsync(msg);
        }

        public static void ConnectTo()
        {
            string serverUrl = "https://nj2020sandbox.crm.dynamics.com";
            string clientId = "f3769397-a593-488f-8233-2048f9ae139e";
            string clientSecret = "__ygP9--d2C.0gX_QhyPcrC22l1XXc-5pf"; // This should be encrypted

            CrmServiceClient sdk = new CrmServiceClient(new Uri(serverUrl), clientId, clientSecret, false, "");
            if (sdk != null && sdk.IsReady)
            {
                sdk.Execute(new WhoAmIRequest());
            }
           var em =  sdk.Retrieve("email", new Guid("aa690ab5-0b30-ea11-a813-000d3a591abb"), new ColumnSet(false));
        }

        public static IOrganizationService GetAccessToken()
        {
            string organizationUrl = "https://nj2020sandbox.crm.dynamics.com";
            string resourceURL = "https://nj2020sandbox.api.crm.dynamics.com" + "/api/data/";
            string clientId = "f3769397-a593-488f-8233-2048f9ae139e"; // Client Id
            string appKey = "__ygP9--d2C.0gX_QhyPcrC22l1XXc-5pf"; //Client Secret

            //Create the Client credentials to pass for authentication
            ClientCredential clientcred = new ClientCredential(clientId, appKey);

            //get the authentication parameters
            //AuthenticationParameters authParam = AuthenticationParameters.CreateFromUrlAsync(new Uri(resourceURL)).Result;

            //Generate the authentication context - this is the azure login url specific to the tenant
            //string authority = authParam.Authority;
            string authority = "https://login.microsoftonline.com/46806687-f2d4-4073-bee0-a5e00bc79c27";
            //request token
            AuthenticationResult authenticationResult = new AuthenticationContext(authority).AcquireTokenAsync(organizationUrl, clientcred).Result;

            //get the token              
            string token = authenticationResult.AccessToken;

            Uri serviceUrl = new Uri(organizationUrl + @"/xrmservices/2011/organization.svc/web?SdkClientVersion=9.0");

            using (var sdkService = new OrganizationWebProxyClient(serviceUrl, false))
            {
                sdkService.HeaderToken = token;

                var _orgService = (IOrganizationService)sdkService != null ? (IOrganizationService)sdkService : null;
                return _orgService;
            }
        }

        public static CrmServiceClient GetService(string url, string login, string password)
        {        
            var connectionString = $"AuthType = OAuth;  Username = {login}; Password = {password};  Url = {url.Trim('/')};" +
                                   "AppId=f3769397-a593-488f-8233-2048f9ae139e;" +
                                   "RedirectUri=app://f3769397-a593-488f-8233-2048f9ae139e;" +
                                   "LoginPrompt=Auto";

            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            return new CrmServiceClient(connectionString);
        }

        public static void CompareEmails(IOrganizationService service)
        {

            /*var emails = GetMails(service);
            var count2 = 0;
            var count3 = 0; 

            foreach(var mail in emails.Entities)
            {
                var size = ASCIIEncoding.Unicode.GetByteCount(mail["description"].ToString());//.Length;
                if (size > 20000) count2++;
            }*/

           // Console.WriteLine("Count more than 2000 - " + count2);*/

            var email1 = GetEmailById(new Guid("9fabd887-ea75-eb11-b1ab-000d3a591fb8"), service);

            var size = ASCIIEncoding.Unicode.GetByteCount(email1["description"].ToString());
            var email5 = GetEmailById(new Guid("e144d6eb-c085-eb11-b1ad-000d3a591e88"), service);
            var size2 = ASCIIEncoding.Unicode.GetByteCount(email5["description"].ToString());

            var email2 = GetEmailById(new Guid("96172e14-c4ad-eb11-8236-000d3a319544"), service);
            var email3 = GetEmailById(new Guid("bfb1a659-c4ad-eb11-8236-000d3a3197c9"), service);
            var email4 = GetEmailById(new Guid("04b18dba-98bc-eb11-8236-000d3a31c36d"), service);

        }

        private static Entity GetEmailById(Guid id, IOrganizationService service)
        {
            return service.Retrieve("email", id, new Microsoft.Xrm.Sdk.Query.ColumnSet("subject", "description", "messageid"));
        }

        public static void DeleteEmails(IOrganizationService service)
        {
            var emailsToDelete = GetMails(service);

            Console.WriteLine("Total - " + emailsToDelete.Count);
            var total = 0;
            foreach(var email in emailsToDelete)
            {
                service.Delete(email.LogicalName, email.Id);
                total++;
                Console.WriteLine($"{total} of {emailsToDelete.Count}");
            }
        }
        public static void ProcessActivityParty(IOrganizationService service)
        {
            var parties = GetParties(service);


        }

        public static List<Entity> GetParties(IOrganizationService service)
        {
            var query = new QueryExpression("activityparty");

            query.ColumnSet = new ColumnSet(true);
            query.Criteria.AddCondition("activityid", ConditionOperator.Equal, new Guid("25c53ebc-a5ce-eb11-8235-000d3a3abd9e"));
            query.Criteria.AddCondition("participationtypemask", ConditionOperator.Equal, 2);
            return service.RetrieveMultiple(query).Entities.ToList();
        }

        public static List<Entity> GetMails(IOrganizationService service)
        {
            var query = new QueryExpression("email");

            query.ColumnSet = new ColumnSet("activityid","messageid");
            query.Criteria.AddCondition("directioncode", ConditionOperator.Equal, false);
           query.Criteria.AddCondition("messageid", ConditionOperator.Like, "%eu.messagegears.net%");
            
            // Set initial page number
            int pageNumber = 1;
            // Collections for entities
            var entityCollection = new EntityCollection();
            var tempResult = new EntityCollection();

            tempResult = service.RetrieveMultiple(query);
            entityCollection.Entities.AddRange(tempResult.Entities);

            // Select records using paging
            do
            {
                Console.WriteLine("Page - " + pageNumber + " T:" + entityCollection.Entities.Count);
                query.PageInfo.Count = 5000;
                query.PageInfo.PageNumber = pageNumber++;
                query.PageInfo.PagingCookie = tempResult.PagingCookie;

                tempResult = service.RetrieveMultiple(query);
                entityCollection.Entities.AddRange(tempResult.Entities);

             
            }
            while (tempResult.MoreRecords);

            return entityCollection.Entities.ToList();
        }
    }
}
