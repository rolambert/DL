using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;


namespace DNLConsole365.Projects
{
    public class TwilloServices
    {
        public static void SendSmsMessage()
        {
            var accountSid = "ACe468f176381eb74a098fd492ffca2e70";
            var authToken = "fbe84df14426b8b40f55e019dcf43397";

            var MessageApiString = string.Format("https://api.twilio.com/2010-04-01/Accounts/{0}/Messages.json", accountSid);

            var request = WebRequest.Create(MessageApiString);
            request.Method = "POST";
            request.Credentials = new NetworkCredential(accountSid, authToken);
            request.ContentType = "application/x-www-form-urlencoded";

            var body = string.Format("From=+18316619752&To=+380962979705&Body=Hello!&{0}:{1}", accountSid, authToken);
            var data = System.Text.ASCIIEncoding.Default.GetBytes(body);

            using (Stream s = request.GetRequestStream())
            {
                s.Write(data, 0, data.Length);
            }

            try
            {
                var result = request.GetResponse();
            }
            catch(Exception ex)
            {
                Console.WriteLine("{0}", ex.Message);
            }
        }

        public static void CheckSharepoint()
        {
            using (var client = new HttpClient {  BaseAddress = new Uri("https://nibus.sharepoint.com") })
            {
                var res = client.GetAsync("/_layouts/15/WopiFrame.aspx?sourcedoc={587e3482-734c-4c4f-8d1b-87630709e3a3}&action=interactivepreview").Result;
                
                Console.WriteLine(res.StatusCode);
            }
        }

        public static void SendSMSNew()
        {
            try
            {
                using (var client = new HttpClient { BaseAddress = new Uri("https://api.twilio.com") })
                {
                    client.DefaultRequestHeaders.Authorization =
                        new AuthenticationHeaderValue("Basic",
                        Convert.ToBase64String(Encoding.ASCII.GetBytes("AC8cabab714f98c11a73c353d6f4e76d51:dee5ebf2865fd6a56f5fe1f4d72d24ab")));

                    var content = new FormUrlEncodedContent(new[]
                    {
                        new KeyValuePair<string, string>("To","+380962979705"),
                        new KeyValuePair<string, string>("From", "+18316619752"),
                        new KeyValuePair<string, string>("Body", "Hello Denis This is cool!")
                     });
                                      
                    var res = client.PostAsync("/2010-04-01/Accounts/AC8cabab714f98c11a73c353d6f4e76d51/Messages.json", content).Result;
                    var b = res.StatusCode;
                    var g = res.IsSuccessStatusCode;
                    var b1 = res.ReasonPhrase;
                    Console.WriteLine();
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine("{0}", ex.Message);
            }
        }
    }
}
