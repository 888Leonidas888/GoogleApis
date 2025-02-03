using System;
using System.Runtime.InteropServices;
using GoogleApis.Sheet.Interfaces;
using GoogleApis.Sheet.Common;

namespace GoogleApis.Sheet.Components
{
    [ComVisible(true)]
    [Guid("38ffbe19-2ef4-484d-9c8b-e742e56ec5ee")]
    [ClassInterface(ClassInterfaceType.None)]
    public class SpreadSheets : GoogleSheetsBase, ISpreadSheets
    {
        public SpreadSheets(string apiKey, string accessToken)
        {
            _apiKey = apiKey;
            _accessToken = accessToken;
        }
        public int Operation()
        {
            return _status;
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
}
