using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace DNLConsole365.Sharepoint
{
    public class SPService : ISharePointService
    {
        private string _username;
        private string _password;
        private string _siteUrl;
        private SpoAuthUtility _spo;
        public SPService(string siteUrl, string username, string password)
        {
            _siteUrl = siteUrl;
            _username = username;
            _password = password;
        }
        

        public SharePointOutput UpdateList(string listName, string id, Dictionary<string, string> fields)
        {
            SharePointOutput result = new SharePointOutput();

            try
            {
                var listItemFullName = GetLIFullName(listName);
                if (!listItemFullName.Success)
                {
                    throw listItemFullName.Exception;
                }
                if (_spo == null)
                    _spo = SpoAuthUtility.Create(new Uri(_siteUrl), _username, WebUtility.HtmlEncode(_password), false);

                string odataQuery = $"/_api/web/lists/GetByTitle('{listName}')/items({id})";

                var jsonDataString = string.Format("{{ '__metadata': {{ 'type': '{1}' }}, {0} }}", ConvertDictionaryToJsonFormat(fields), listItemFullName.SPRestData.d.ListItemEntityTypeFullName);

                byte[] content = ASCIIEncoding.ASCII.GetBytes(jsonDataString);

                string digest = _spo.GetRequestDigest();

                Uri url = new Uri(String.Format("{0}/{1}", _spo.SiteUrl, odataQuery));

                // Set X-RequestDigest
                var webRequest = (HttpWebRequest)HttpWebRequest.Create(url);
                webRequest.Headers.Add("X-RequestDigest", digest);
                webRequest.Headers.Add("X-HTTP-Method", "MERGE");
                webRequest.Headers.Add("IF-MATCH", "*");


                // Send a json odata request to SPO rest services to fetch all list items for the list.
                var response = HttpHelper.SendODataJsonRequest(url, "POST", content, webRequest, _spo);
                if (!response.Success)
                {
                    result.Success = false;
                    result.Exception = new Exception($"Failed to process the action on Sharepoint. HTTP Staus Code - {response.Status} ");
                }
                else
                {
                    var responseJson = Encoding.UTF8.GetString(response.Response, 0, response.Response.Length);
                    
                   // Response from Sharepoint REST API is 0 byte length (surprise :O)
                   // So no meaning of converting any thing from here..
                   // If you want the full list item to return, do a service call again over here and return the obj..
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Exception = ex;
            }
            return result;
        }

        public SharePointOutput GetFilesFromFolder(string folderName)
        {
            SharePointOutput result = new SharePointOutput();

            if (_spo == null)
                _spo = SpoAuthUtility.Create(new Uri(_siteUrl), _username, WebUtility.HtmlEncode(_password), false);

            string odataQuery = $"_api/web/getfolderbyserverrelativeurl('{folderName}')/files";
                                
            string digest = _spo.GetRequestDigest();

            Uri url = new Uri(String.Format("{0}/{1}", _spo.SiteUrl, odataQuery));

            // Set X-RequestDigest
            var webRequest = (HttpWebRequest)HttpWebRequest.Create(url);
            webRequest.Headers.Add("X-RequestDigest", digest);

            // Send a json odata request to SPO rest services to fetch all list items for the list.
            var response = HttpHelper.SendODataJsonRequest(url, "GET", null, webRequest, _spo);
            if (!response.Success)
            {
                result.Success = false;
                result.Exception = new Exception($"Failed to process the action on Sharepoint. HTTP Staus Code - {response.Status} ");
            }
            else
            {
                var responseJson = Encoding.UTF8.GetString(response.Response, 0, response.Response.Length);
                SharepointRestData spObj = null;
                try
                {
                    spObj = Searilizer.DeserilizeSPJsonString(responseJson);
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.Exception = new Exception("Failed to deserialize Sharepoint REST Json string. Please see Inner Exception", ex);
                    return result;
                }
                result.SPRestData = spObj;
                result.JsonRestData = responseJson;
            }
            return result;

        }


        public SharePointOutput CreateFolderForOpportunity(string folderName)
        {
            SharePointOutput result = new SharePointOutput();

            if (_spo == null)
                _spo = SpoAuthUtility.Create(new Uri(_siteUrl), _username, WebUtility.HtmlEncode(_password), false);

            
            string odataQuery = $"/_api/Web/Folders";
            string digest = _spo.GetRequestDigest();

            var url = new Uri(String.Format("{0}/{1}", _spo.SiteUrl, odataQuery));
            
            byte[] content = ASCIIEncoding.ASCII.GetBytes("{'__metadata': { 'type': 'SP.Folder' }, 'ServerRelativeUrl': 'Opportunity/" + folderName + "'}");

            // Set X-RequestDigest
            var webRequest = (HttpWebRequest)HttpWebRequest.Create(url);
            webRequest.Headers.Add("X-RequestDigest", digest);

            // Send a json odata request to SPO rest services to fetch all list items for the list.
            var response = HttpHelper.SendODataJsonRequest(url, "POST", content, webRequest, _spo);
            if (!response.Success)
            {
                result.Success = false;
                result.Exception = new Exception($"Failed to process the action on Sharepoint. HTTP Staus Code - {response.Status} ");
            }
            else
            {
                var responseJson = Encoding.UTF8.GetString(response.Response, 0, response.Response.Length);
                SharepointRestData spObj = null;
                try
                {
                    spObj = Searilizer.DeserilizeSPJsonString(responseJson);
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.Exception = new Exception("Failed to deserialize Sharepoint REST Json string. Please see Inner Exception", ex);
                    return result;
                }
                result.SPRestData = spObj;
                result.JsonRestData = responseJson;
            }
            return result;            
        }

        public SharePointOutput GetItemsFromFolder(string folderName)
        {
            SharePointOutput result = new SharePointOutput();

            if (_spo == null)
                _spo = SpoAuthUtility.Create(new Uri(_siteUrl), _username, WebUtility.HtmlEncode(_password), false);

            string odataQuery = $"_api/web/getfolderbyserverrelativeurl('{folderName}')?$expand=Folders,Files";

            string digest = _spo.GetRequestDigest();

            Uri url = new Uri(String.Format("{0}/{1}", _spo.SiteUrl, odataQuery));

            // Set X-RequestDigest
            var webRequest = (HttpWebRequest)HttpWebRequest.Create(url);
            webRequest.Headers.Add("X-RequestDigest", digest);

            // Send a json odata request to SPO rest services to fetch all list items for the list.
            var response = HttpHelper.SendODataJsonRequest(url, "GET", null, webRequest, _spo);
            if (!response.Success)
            {
                result.Success = false;
                result.Exception = new Exception($"Failed to process the action on Sharepoint. HTTP Staus Code - {response.Status} ");
            }
            else
            {
                var responseJson = Encoding.UTF8.GetString(response.Response, 0, response.Response.Length);
                SharepointRestData spObj = null;
                try
                {
                    spObj = Searilizer.DeserilizeSPJsonString(responseJson);
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.Exception = new Exception("Failed to deserialize Sharepoint REST Json string. Please see Inner Exception", ex);
                    return result;
                }
                result.SPRestData = spObj;
                result.JsonRestData = responseJson;
            }
            return result;

        }

        public SharePointOutput GetFoldersFromFolder(string folderName)
        {
            SharePointOutput result = new SharePointOutput();

            if (_spo == null)
                _spo = SpoAuthUtility.Create(new Uri(_siteUrl), _username, WebUtility.HtmlEncode(_password), false);

            string odataQuery = $"_api/web/getfolderbyserverrelativeurl('{folderName}')/folders";

            string digest = _spo.GetRequestDigest();

            Uri url = new Uri(String.Format("{0}/{1}", _spo.SiteUrl, odataQuery));

            // Set X-RequestDigest
            var webRequest = (HttpWebRequest)HttpWebRequest.Create(url);
            webRequest.Headers.Add("X-RequestDigest", digest);

            // Send a json odata request to SPO rest services to fetch all list items for the list.
            var response = HttpHelper.SendODataJsonRequest(url, "GET", null, webRequest, _spo);
            if (!response.Success)
            {
                result.Success = false;
                result.Exception = new Exception($"Failed to process the action on Sharepoint. HTTP Staus Code - {response.Status} ");
            }
            else
            {
                var responseJson = Encoding.UTF8.GetString(response.Response, 0, response.Response.Length);
                SharepointRestData spObj = null;
                try
                {
                    spObj = Searilizer.DeserilizeSPJsonString(responseJson);
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.Exception = new Exception("Failed to deserialize Sharepoint REST Json string. Please see Inner Exception", ex);
                    return result;
                }
                result.SPRestData = spObj;
                result.JsonRestData = responseJson;
            }
            return result;

        }

        public SharePointOutput GetFolderItems(string libraryName)
        {
            SharePointOutput result = new SharePointOutput();

            if (_spo == null)
                _spo = SpoAuthUtility.Create(new Uri(_siteUrl), _username, WebUtility.HtmlEncode(_password), false);

            //string odataQuery = $"_api/web/lists/GetByTitle('Opportunity')";
            string odataQuery = $"_api/web/lists/GetByTitle('{libraryName}')/rootFolder/Folders";

            //string odataQuery = $"_api/web/lists/GetByTitle('{listName}')";

            string digest = _spo.GetRequestDigest();

            Uri url = new Uri(String.Format("{0}/{1}", _spo.SiteUrl, odataQuery));

            // Set X-RequestDigest
            var webRequest = (HttpWebRequest)HttpWebRequest.Create(url);
            webRequest.Headers.Add("X-RequestDigest", digest);

            // Send a json odata request to SPO rest services to fetch all list items for the list.
            var response = HttpHelper.SendODataJsonRequest(url, "GET", null, webRequest, _spo);
            if (!response.Success)
            {
                result.Success = false;
                result.Exception = new Exception($"Failed to process the action on Sharepoint. HTTP Staus Code - {response.Status} ");
            }
            else
            {
                var responseJson = Encoding.UTF8.GetString(response.Response, 0, response.Response.Length);
                SharepointRestData spObj = null;
                try
                {
                    spObj = Searilizer.DeserilizeSPJsonString(responseJson);
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.Exception = new Exception("Failed to deserialize Sharepoint REST Json string. Please see Inner Exception", ex);
                    return result;
                }
                result.SPRestData = spObj;
                result.JsonRestData = responseJson;
            }
            return result;

        }

        public SharePointOutput GetListItems(string libraryName)
        {
            SharePointOutput result = new SharePointOutput();

            if (_spo == null)
                _spo = SpoAuthUtility.Create(new Uri(_siteUrl), _username, WebUtility.HtmlEncode(_password), false);

            string odataQuery = $"_api/Web/Lists/getbytitle('{libraryName}')/items";

            //string odataQuery = $"_api/web/lists/GetByTitle('{listName}')";

            string digest = _spo.GetRequestDigest();

            Uri url = new Uri(String.Format("{0}/{1}", _spo.SiteUrl, odataQuery));

            // Set X-RequestDigest
            var webRequest = (HttpWebRequest)HttpWebRequest.Create(url);
            webRequest.Headers.Add("X-RequestDigest", digest);

            // Send a json odata request to SPO rest services to fetch all list items for the list.
            var response = HttpHelper.SendODataJsonRequest(url, "GET", null, webRequest, _spo);
            if (!response.Success)
            {
                result.Success = false;
                result.Exception = new Exception($"Failed to process the action on Sharepoint. HTTP Staus Code - {response.Status} ");
            }
            else
            {
                var responseJson = Encoding.UTF8.GetString(response.Response, 0, response.Response.Length);
                SharepointRestData spObj = null;
                try
                {
                    spObj = Searilizer.DeserilizeSPJsonString(responseJson);
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.Exception = new Exception("Failed to deserialize Sharepoint REST Json string. Please see Inner Exception", ex);
                    return result;
                }
                result.SPRestData = spObj;
            }
            return result;
            
        }


        public SharePointOutput GetLIFullName(string listName)
        {
            SharePointOutput result = new SharePointOutput();

            if (_spo==null)
                _spo = SpoAuthUtility.Create(new Uri(_siteUrl), _username, WebUtility.HtmlEncode(_password), false);

            string odataQuery = $"_api/web/lists/GetByTitle('{listName}')";

            string digest = _spo.GetRequestDigest();

            Uri url = new Uri(String.Format("{0}/{1}", _spo.SiteUrl, odataQuery));

            // Set X-RequestDigest
            var webRequest = (HttpWebRequest)HttpWebRequest.Create(url);
            webRequest.Headers.Add("X-RequestDigest", digest);

            // Send a json odata request to SPO rest services to fetch all list items for the list.
            var response = HttpHelper.SendODataJsonRequest(url, "GET", null, webRequest, _spo);
            if (!response.Success)
            {
                result.Success = false;
                result.Exception = new Exception($"Failed to process the action on Sharepoint. HTTP Staus Code - {response.Status} ");
            }
            else
            {
                var responseJson = Encoding.UTF8.GetString(response.Response, 0, response.Response.Length);
                SharepointRestData spObj = null;
                try
                {
                    spObj = Searilizer.DeserilizeSPJsonString(responseJson);
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.Exception = new Exception("Failed to deserialize Sharepoint REST Json string. Please see Inner Exception", ex);
                    return result;
                }
                result.SPRestData = spObj;
            }
            return result;
        }

        public SharePointOutput CreateListItem(string listName, Dictionary<string, string> fields)
        {
            SharePointOutput result = new SharePointOutput();

            try
            {
                var listItemFullName = GetLIFullName(listName);
                if (!listItemFullName.Success)
                {
                    throw listItemFullName.Exception;
                }
                if (_spo == null)
                    _spo = SpoAuthUtility.Create(new Uri(_siteUrl), _username, WebUtility.HtmlEncode(_password), false);

                //string odataQuery = $"/_api/web/lists/GetByTitle('Leads from Flow')/items({id})";
                string odataQuery = $"_api/web/lists/GetByTitle('{listName}')/items";
                
                var jsonDataString = string.Format("{{ '__metadata': {{ 'type': '{1}' }}, {0} }}", ConvertDictionaryToJsonFormat(fields), listItemFullName.SPRestData.d.ListItemEntityTypeFullName);

                byte[] content = ASCIIEncoding.ASCII.GetBytes(jsonDataString);

                string digest = _spo.GetRequestDigest();

                Uri url = new Uri(String.Format("{0}/{1}", _spo.SiteUrl, odataQuery));
                
                // Set X-RequestDigest
                var webRequest = (HttpWebRequest)HttpWebRequest.Create(url);
                webRequest.Headers.Add("X-RequestDigest", digest);

                // Send a json odata request to SPO rest services to fetch all list items for the list.
                var response = HttpHelper.SendODataJsonRequest(url, "POST", content, webRequest, _spo);
                if (!response.Success)
                {
                    result.Success = false;
                    result.Exception = new Exception($"Failed to process the action on Sharepoint. HTTP Staus Code - {response.Status} ");
                }
                else
                {
                    var responseJson = Encoding.UTF8.GetString(response.Response, 0, response.Response.Length);
                    
                    SharepointRestData spObj = null;
                    try
                    {
                        spObj = Searilizer.DeserilizeSPJsonString(responseJson);
                        
                    }
                    catch (Exception ex)
                    {
                        result.Success = false;
                        result.Exception = new Exception("Failed to deserialize Sharepoint REST Json string. Please see Inner Exception", ex);
                        return result;
                    }
                    result.SPRestData = spObj;
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Exception = ex;
            }
            return result;
        }

        private string ConvertDictionaryToJsonFormat(Dictionary<string, string> data)
        {
            //, 'Title': 'Test'
            var result = string.Empty;

            if (data == null)
                return result;
            foreach (var item in data)
            {
                result = string.IsNullOrEmpty(result) ? $"'{item.Key}':'{item.Value}'" : $"{result},'{item.Key}':'{item.Value}'";
            }

            return result;
        }

        
    }


}
