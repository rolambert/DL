using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
using System.Net;
using System.Net.Http.Headers;

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using Stripe;

namespace DNLConsole365.Projects
{
    
    public class StripeServices
    {
        private const string testSecretKey = "sk_test_Z3RxLEb8ePEP9jtfLdsoIWJO";

        public static void CreateCustomer(string email)
        {
            StripeConfiguration.SetApiKey(testSecretKey);

            var options = new CustomerCreateOptions
            {
                Description = "Some Test Description",
                Email = email                
            };

            var service = new CustomerService();
            Customer customer = service.Create(options);
        }

        public static DateTime UnixTimeStampToDateTime(long unixTime)
        {
            var unixStart = new DateTime(1970, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc);
            var unixTimeStampInTicks = (unixTime * TimeSpan.TicksPerSecond);

            return new DateTime(unixStart.Ticks + unixTimeStampInTicks, DateTimeKind.Utc);
            //return new DateTime(unixStart.Ticks + unixTimeStampInTicks);
        }

        public static long DateTimeToUnixTimestamp(DateTime dateTime)
        {
            var unixStart = new DateTime(1970, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc);
            var unixTimeStampInTicks = (dateTime.ToUniversalTime() - unixStart).Ticks;

            return (unixTimeStampInTicks / TimeSpan.TicksPerSecond);
        }

        public static string ExecuteStripeGetRequest(string url)
        {

            using (var client = new HttpClient())
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;

                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", testSecretKey);

                //var response = client.PostAsync(url, null).Result;
                var response = client.GetAsync(url).Result;
                var result = response.Content.ReadAsStringAsync().Result;

                return result;
            }
        }

        public static List<Entity> GetCrmPayments(IOrganizationService service, DateTime date)
        {
            var query = new QueryExpression("dnl_paymententity");
            query.ColumnSet = new ColumnSet("dnl_stripeid", "dnl_customer");
            query.Criteria.AddCondition("createdon", ConditionOperator.LastXDays, 1);

            return service.RetrieveMultiple(query).Entities.ToList();
        }
    }
}
