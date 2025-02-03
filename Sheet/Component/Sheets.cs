using System;
using System.Runtime.InteropServices;
using GoogleApis.Sheet.Interfaces;
using GoogleApis.Sheet.Common;

namespace GoogleApis.Sheet.Components
{
    [ComVisible(true)]
    [Guid("c7c37020-6701-4496-80c2-2f25b3e15ce8")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Sheets : GoogleSheetsBase, ISheets
    {
        public Sheets(string apiKey, string accessToken)
        {
            _apiKey = apiKey;
            _accessToken = accessToken;
        }
        public int Operation()
        {
            return _status;
        }
        public string CopyTo(string spreadsheetId, string sheetId, string json = null, Scripting.Dictionary queryParameters = null)
        {
            try
            {
                string pathParameters = $"/v4/spreadsheets/{spreadsheetId}/sheets/{sheetId}:copyTo";

                string url = CreateQueryParameters(pathParameters, queryParameters);

                var response = Request("POST", url, json);
                return response.Content.ReadAsStringAsync().Result;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error CopyTo : {ex.Message}");
            }
        }
    }
}
