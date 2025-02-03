using System;
using System.Runtime.InteropServices;
using System.Net.Http;
using System.IO;
using Newtonsoft.Json.Linq;
using System.Diagnostics;
using Newtonsoft.Json;
using System.Net.Http.Headers;
using System.Text;
using GoogleApis.Drive.Helper;
using GoogleApis.Drive.Interfaces;
using GoogleApis.Drive.Component;

namespace GoogleApis
{    
    [ComVisible(true)]
    [Guid("9f8b1b7e-0c42-40e3-89a7-4bda8a8e2e30")]
    [ClassInterface(ClassInterfaceType.None)]
    public class GoogleDrive: IGoogleDrive
    {
        private const string SERVICE_END_POINT = "https://www.googleapis.com/drive/v3/files";
        private const string SERVICE_END_POINT_UPLOAD = "https://www.googleapis.com/upload/drive/v3/files";

        private string _apiKey;
        private string _accessToken;
        private int _status;        

        public int Operation()
        {
            return _status;
        }
        public void ConnectionService(object oFlowOauth)
        {
            dynamic flow = oFlowOauth;

            _apiKey = flow.GetApiKey();
            _accessToken = flow.GetTokenAccess();
        }
        private string CreateQueryParameters(string pathParameters = null, Scripting.Dictionary queryParameters = null, bool endPointUpload = false)
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
                            var v = Helper.URLEncode(dict.get_Item(k));
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
        private HttpResponseMessage Request(string method, string url, object body = null, Scripting.Dictionary headers = null)
        {
            if (headers == null)
            {
                headers = new Scripting.Dictionary();
            }

            headers.set_Item("Authorization", $"Bearer {_accessToken}");
            headers.set_Item("Accept", "application/json");

            var response = Helper.Request(method, url, body, headers);
            _status = (int)response.StatusCode;
            return response;
        }
        public Files Files()
        {
            return new Files(this._apiKey, this._accessToken);
        }

    }
}
