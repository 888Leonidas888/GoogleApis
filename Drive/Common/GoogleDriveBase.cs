using System;
using System.Net.Http;

namespace GoogleApis.Drive.Common
{
    public class GoogleDriveBase
    {
        protected const string SERVICE_END_POINT = "https://www.googleapis.com/drive/v3/files";
        protected const string SERVICE_END_POINT_UPLOAD = "https://www.googleapis.com/upload/drive/v3/files";

        protected string _apiKey;
        protected string _accessToken;
        protected int _status;
        protected string CreateQueryParameters(string pathParameters = null, Scripting.Dictionary queryParameters = null, bool endPointUpload = false)
        {
            string endPoint = endPointUpload ? SERVICE_END_POINT_UPLOAD : SERVICE_END_POINT;
            string queryString = "?";

            try
            {
                if (queryParameters != null)
                {
                    if (queryParameters is Scripting.Dictionary dict)
                    {
                        foreach (var k in dict)
                        {
                            var v = Helper.Helper.URLEncode(dict.get_Item(k));
                            queryString += $"{k}={v}&";
                        }
                    }
                    else
                    {
                        throw new Exception($"QueryParameters not is dicctionary.");
                    }
                }

                endPoint += pathParameters;
                queryString += $"key={_apiKey}";

                return $"{endPoint}{queryString}";
            }
            catch (Exception ex)
            {
                throw new Exception($"Error create query parameters: {ex.Message}");
            }
        }
        protected HttpResponseMessage Request(string method, string url, object body = null, Scripting.Dictionary headers = null)
        {
            if (headers == null)
            {
                headers = new Scripting.Dictionary();
            }

            headers.set_Item("Authorization", $"Bearer {_accessToken}");
            headers.set_Item("Accept", "application/json");

            var response = Helper.Helper.Request(method, url, body, headers);
            _status = (int)response.StatusCode;
            return response;
        }        
    }
}
