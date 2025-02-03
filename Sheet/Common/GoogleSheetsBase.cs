using System;
using System.Net.Http;

namespace GoogleApis.Sheet.Common
{
    public class GoogleSheetsBase
    {
        protected const string SERVICE_END_POINT = "https://sheets.googleapis.com";
        protected string _apiKey;
        protected string _accessToken;
        protected int _status;

        protected string CreateQueryParameters(string pathParameters = null, Scripting.Dictionary queryParameters = null)
        {
            string endPoint = SERVICE_END_POINT;
            string queryString = "?";

            try
            {
                if (queryParameters != null)
                {
                    foreach (var k in queryParameters)
                    {
                        var v = Helper.Helper.URLEncode(queryParameters.get_Item(k));
                        queryString += $"{k}={v}&";
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
            headers.set_Item("Content-Type", "application/json");

            var response = Helper.Helper.Request(method, url, body, headers);
            _status = (int)response.StatusCode;
            return response;
        }
    }
}
