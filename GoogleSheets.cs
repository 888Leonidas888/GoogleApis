using System;
using System.Runtime.InteropServices;
using System.Net.Http;

namespace GoogleApis
{
    [ComVisible(true)]
    [Guid("62e6123b-bd17-479d-9276-87536b16c584")]
    public interface IGoogleSheets
    {
        int Operation();
        void ConnectionService(object oFlowOauth);
        SpreadSheets SpreadSheets();
    }

    [ComVisible(true)]
    [Guid("6c92c2a1-5812-4ebc-9e72-b49431131c1c")]
    public interface ISpreadSheets
    {
        int Operation();
        void ConnectionService(object oFlowOauth);
        string BatchUpdate(string spreadsheetId, string json, Scripting.Dictionary queryParameters = null);
        string Create(string json = null, Scripting.Dictionary queryParameters = null);
        string Get(string spreadsheetId, Scripting.Dictionary queryParameters = null);
        string GetByDataFilter(string spreadsheetId, string json, Scripting.Dictionary queryParameters = null);
        Sheets Sheets();
        Values Values();
    }

    [ComVisible(true)]
    [Guid("89ccf222-d621-4736-8c53-8c347f7502b8")]
    public interface ISheets
    {
        string CopyTo(string spreadsheetId, string sheetId, string json = null, Scripting.Dictionary queryParameters = null);
    }

    [ComVisible(true)]
    [Guid("8c2e3e38-b3c5-4cf4-b7b0-5b963ef60349")]
    public interface IValues
    {
        string Append(string spreadsheetId, string rng, string json, Scripting.Dictionary queryParameters);
        string BatchClear(string spreadsheetId, string json, Scripting.Dictionary queryParameters);
        string BatchClearByDataFilter(string spreadsheetId, string json, Scripting.Dictionary queryParameters);
        string BatchGet(string spreadsheetId, Scripting.Dictionary queryParameters = null);
        string BatchGetByDataFilter(string spreadsheetId, string json, Scripting.Dictionary queryParameters = null);
        string BatchUpdate(string spreadsheetId, string json, Scripting.Dictionary queryParameters = null);
        string BatchUpdateByDataFilter(string spreadsheetId, string json, Scripting.Dictionary queryParameters = null);
        string Clear(string spreadsheetId, string rng, Scripting.Dictionary queryParameters = null);
        string Get(string spreadsheetId, string rng, Scripting.Dictionary queryParameters = null);
        string Update(string spreadsheetId, string rng, string json, Scripting.Dictionary queryParameters = null);
    }

    [ComVisible(true)]
    [Guid("cdd4e85b-c33b-4e9b-b825-845307f9c190")]
    [ClassInterface(ClassInterfaceType.None)]
    public class GoogleSheets: IGoogleSheets
    {
        private const string SERVICE_END_POINT = "https://sheets.googleapis.com";

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
        private string CreateQueryParameters(string pathParameters = null, Scripting.Dictionary queryParameters = null)
        {
            string endPoint = SERVICE_END_POINT;
            string queryString = "?";

            try
            {
                if (queryParameters != null)
                {
                    foreach (var k in queryParameters)
                    {
                        var v = Helper.URLEncode(queryParameters.get_Item(k));
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
        private HttpResponseMessage Request(string method, string url, object body = null, Scripting.Dictionary headers = null)
        {
            if (headers == null)
            {
                headers = new Scripting.Dictionary();
            }

            headers.set_Item("Authorization", $"Bearer {_accessToken}");
            headers.set_Item("Accept", "application/json");
            headers.set_Item("Content-Type", "application/json");

            var response = Helper.Request(method, url, body, headers);
            _status = (int)response.StatusCode;
            return response;
        }
        public SpreadSheets SpreadSheets()
        {
            return new SpreadSheets(this._apiKey,this._accessToken);
        }
    }

    [ComVisible(true)]
    [Guid("38ffbe19-2ef4-484d-9c8b-e742e56ec5ee")]
    [ClassInterface(ClassInterfaceType.None)] 
    public class SpreadSheets : ISpreadSheets
    {
        private const string SERVICE_END_POINT = "https://sheets.googleapis.com";
       
        private string _apiKey;
        private string _accessToken;
        private int _status;        
        
        public SpreadSheets(string apiKey, string accessToken)
        {
            _apiKey = apiKey;
            _accessToken = accessToken;
        }
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
        private string CreateQueryParameters(string pathParameters = null, Scripting.Dictionary queryParameters = null)
        {
            string endPoint = SERVICE_END_POINT;
            string queryString = "?";

            try
            {
                if (queryParameters != null)
                {                          
                    foreach (var k in queryParameters)
                    {
                        var v = Helper.URLEncode(queryParameters.get_Item(k));
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
        private HttpResponseMessage Request(string method, string url, object body = null, Scripting.Dictionary headers = null)
        {
            if (headers == null)
            {
                headers = new Scripting.Dictionary();
            }

            headers.set_Item("Authorization", $"Bearer {_accessToken}");
            headers.set_Item("Accept", "application/json");
            headers.set_Item("Content-Type", "application/json");

            var response =  Helper.Request(method, url, body, headers);
            _status = (int)response.StatusCode;
            return response;
        }        
        public string BatchUpdate(string spreadsheetId, string json, Scripting.Dictionary queryParameters = null)
        {
            try
            {
                string pathParameters = $"/v4/spreadsheets/{spreadsheetId}:batchUpdate";               

                string url = CreateQueryParameters(pathParameters, queryParameters);

                var response = Request("POST", url, json);
                return response.Content.ReadAsStringAsync().Result;                
            }
            catch (Exception ex)
            {
                throw new Exception($"Error BatchUpdate {spreadsheetId} : {ex.Message}");
            }
        }
        public string Create(string json = null, Scripting.Dictionary queryParameters = null) 
        {
            try
            {
                string pathParameters = $"/v4/spreadsheets";                

                string url = CreateQueryParameters(pathParameters, queryParameters);

                var response = Request("POST", url, json);
                return response.Content.ReadAsStringAsync().Result;               
            }
            catch (Exception ex)
            {
                throw new Exception($"Error Create : {ex.Message}");
            }
        }     
        public string Get(string spreadsheetId, Scripting.Dictionary queryParameters = null)
        {
            try
            {
                string pathParameters = $"/v4/spreadsheets/{spreadsheetId}";                

                string url = CreateQueryParameters(pathParameters, queryParameters);

                var response = Request("GET", url);
                return response.Content.ReadAsStringAsync().Result;                
            }
            catch (Exception ex)
            {
                throw new Exception($"Error Get : {ex.Message}");
            }
        }
        public string GetByDataFilter(string spreadsheetId, string json = null, Scripting.Dictionary queryParameters = null)
        {
            try
            {
                string pathParameters = $"/v4/spreadsheets/{spreadsheetId}:getByDataFilter";

                string url = CreateQueryParameters(pathParameters, queryParameters);

                var response = Request("POST", url, json);
                return response.Content.ReadAsStringAsync().Result;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error GetByDataFilter : {ex.Message}");
            }
        }
        public Sheets Sheets()
        {
            return new Sheets(this._apiKey, this._accessToken);
        }
        public Values Values() 
        { 
            return new Values(this._apiKey, this._accessToken);
        }
    }

    [ComVisible(true)]
    [Guid("c7c37020-6701-4496-80c2-2f25b3e15ce8")] 
    [ClassInterface(ClassInterfaceType.None)]
    public class Sheets : ISheets
    {
        private const string SERVICE_END_POINT = "https://sheets.googleapis.com";

        private string _apiKey;
        private string _accessToken;
        private int _status;        

        public Sheets(string apiKey, string accessToken)
        {
            _apiKey = apiKey;
            _accessToken = accessToken;
        }
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
        private string CreateQueryParameters(string pathParameters = null, Scripting.Dictionary queryParameters = null)
        {
            string endPoint = SERVICE_END_POINT;
            string queryString = "?";

            try
            {
                if (queryParameters != null)
                {
                    foreach (var k in queryParameters)
                    {
                        var v = Helper.URLEncode(queryParameters.get_Item(k));
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
        private HttpResponseMessage Request(string method, string url, object body = null, Scripting.Dictionary headers = null)
        {
            if (headers == null)
            {
                headers = new Scripting.Dictionary();
            }

            headers.set_Item("Authorization", $"Bearer {_accessToken}");
            headers.set_Item("Accept", "application/json");
            headers.set_Item("Content-Type", "application/json");

            var response = Helper.Request(method, url, body, headers);
            _status = (int)response.StatusCode;
            return response;
        }
        public string CopyTo(string spreadsheetId, string sheetId, string json = null, Scripting.Dictionary queryParameters = null)
        {
            try
            {
                string pathParameters = $"/v4/spreadsheets/{spreadsheetId}/sheets/{sheetId}:copyTo";

                string url = CreateQueryParameters(pathParameters,queryParameters);

                var response = Request("POST", url, json);
                return response.Content.ReadAsStringAsync().Result;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error CopyTo : {ex.Message}");
            }
        }
    }

    [ComVisible(true)]
    [Guid("478e0e93-ce20-4388-bc54-672b8d68562a")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Values : IValues
    {
        private const string SERVICE_END_POINT = "https://sheets.googleapis.com";

        private string _apiKey;
        private string _accessToken;
        private int _status;        

        public Values(string apiKey, string accessToken)
        {
            _apiKey = apiKey;
            _accessToken = accessToken;           
        }

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
        private string CreateQueryParameters(string pathParameters = null, Scripting.Dictionary queryParameters = null)
        {
            string endPoint = SERVICE_END_POINT;
            string queryString = "?";

            try
            {
                if (queryParameters != null)
                {
                    foreach (var k in queryParameters)
                    {
                        var v = Helper.URLEncode(queryParameters.get_Item(k));
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
        private HttpResponseMessage Request(string method, string url, object body = null, Scripting.Dictionary headers = null)
        {
            if (headers == null)
            {
                headers = new Scripting.Dictionary();
            }

            headers.set_Item("Authorization", $"Bearer {_accessToken}");
            headers.set_Item("Accept", "application/json");
            headers.set_Item("Content-Type", "application/json");

            var response = Helper.Request(method, url, body, headers);
            _status = (int)response.StatusCode;
            return response;
        }
        public string Append(string spreadsheetId, string rng, string json, Scripting.Dictionary queryParameters)
        {
            try
            {
                string pathParameters = $"/v4/spreadsheets/{spreadsheetId}/values/{rng}:append";

                string url = CreateQueryParameters(pathParameters, queryParameters);

                var response = Request("POST", url, json);
                return response.Content.ReadAsStringAsync().Result;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error Append : {ex.Message}");
            }
        }
        public string BatchClear(string spreadsheetId, string json, Scripting.Dictionary queryParameters)
        {
            try
            {
                string pathParameters = $"/v4/spreadsheets/{spreadsheetId}/values:batchClear";

                string url = CreateQueryParameters(pathParameters, queryParameters);

                var response = Request("POST", url, json);
                return response.Content.ReadAsStringAsync().Result;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error BatchClear : {ex.Message}");
            }
        }
        public string BatchClearByDataFilter(string spreadsheetId, string json, Scripting.Dictionary queryParameters)
        {
            try
            {
                string pathParameters = $"/v4/spreadsheets/{spreadsheetId}/values:batchClearByDataFilter";

                string url = CreateQueryParameters(pathParameters, queryParameters);

                var response = Request("POST", url, json);
                return response.Content.ReadAsStringAsync().Result;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error BatchClearByDataFilter : {ex.Message}");
            }
        }
        public string BatchGet(string spreadsheetId, Scripting.Dictionary queryParameters = null)
        {
            try
            {
                string pathParameters = $"/v4/spreadsheets/{spreadsheetId}/values:batchGet";

                string url = CreateQueryParameters(pathParameters, queryParameters);

                var response = Request("GET", url);
                return response.Content.ReadAsStringAsync().Result;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error BatchGet : {ex.Message}");
            }
        }
        public string BatchGetByDataFilter(string spreadsheetId, string json, Scripting.Dictionary queryParameters = null)
        {
            try
            {
                string pathParameters = $"/v4/spreadsheets/{spreadsheetId}/values:batchGetByDataFilter";

                string url = CreateQueryParameters(pathParameters, queryParameters);

                var response = Request("POST", url, json);
                return response.Content.ReadAsStringAsync().Result;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error BatchGetByDataFilter : {ex.Message}");
            }
        }
        public string BatchUpdate(string spreadsheetId, string json, Scripting.Dictionary queryParameters = null)
        {
            try
            {
                string pathParameters = $"/v4/spreadsheets/{spreadsheetId}/values:batchUpdate";

                string url = CreateQueryParameters(pathParameters, queryParameters);

                var response = Request("POST", url, json);
                return response.Content.ReadAsStringAsync().Result;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error BatchUpdate : {ex.Message}");
            }
        }
        public string BatchUpdateByDataFilter(string spreadsheetId, string json, Scripting.Dictionary queryParameters = null)
        {
            try
            {
                string pathParameters = $"/v4/spreadsheets/{spreadsheetId}/values:batchUpdateByDataFilter";

                string url = CreateQueryParameters(pathParameters, queryParameters);

                var response = Request("POST", url, json);
                return response.Content.ReadAsStringAsync().Result;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error BatchUpdateByDataFilter : {ex.Message}");
            }
        }
        public string Clear(string spreadsheetId, string rng, Scripting.Dictionary queryParameters = null)
        {
            try
            {
                string pathParameters = $"/v4/spreadsheets/{spreadsheetId}/values/{rng}:clear";

                string url = CreateQueryParameters(pathParameters, queryParameters);

                var response = Request("POST", url);
                return response.Content.ReadAsStringAsync().Result;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error Clear : {ex.Message}");
            }
        }
        public string Get(string spreadsheetId, string rng, Scripting.Dictionary queryParameters = null)
        {
            try
            {
                string pathParameters = $"/v4/spreadsheets/{spreadsheetId}/values/{rng}";

                string url = CreateQueryParameters(pathParameters, queryParameters);

                var response = Request("GET", url);
                return response.Content.ReadAsStringAsync().Result;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error Get : {ex.Message}");
            }
        }
        public string Update(string spreadsheetId, string rng, string json, Scripting.Dictionary queryParameters = null)
        {
            try
            {
                string pathParameters = $"/v4/spreadsheets/{spreadsheetId}/values/{rng}";

                string url = CreateQueryParameters(pathParameters, queryParameters);

                var response = Request("PUT", url, json);
                return response.Content.ReadAsStringAsync().Result;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error Update : {ex.Message}");
            }
        }
    }
}
