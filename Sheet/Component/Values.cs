using System;
using System.Runtime.InteropServices;
using GoogleApis.Sheet.Interfaces;
using GoogleApis.Sheet.Common;

namespace GoogleApis.Sheet.Components
{
    [ComVisible(true)]
    [Guid("478e0e93-ce20-4388-bc54-672b8d68562a")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Values : GoogleSheetsBase, IValues
    {
        public Values(string apiKey, string accessToken)
        {
            _apiKey = apiKey;
            _accessToken = accessToken;
        }
        public int Operation()
        {
            return _status;
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
