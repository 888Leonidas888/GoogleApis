using System;
using System.Runtime.InteropServices;
using GoogleApis.Sheet.Interfaces;
using GoogleApis.Sheet.Common;

namespace GoogleApis.Sheet.Components
{
    /// <summary>
    /// Represents the values component of the Google Sheets API.
    /// </summary>
    [ComVisible(true)]
    [Guid("478e0e93-ce20-4388-bc54-672b8d68562a")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Values : GoogleSheetsBase, IValues
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="Values"/> class.
        /// </summary>
        /// <param name="apiKey">The API key.</param>
        /// <param name="accessToken">The access token.</param>
        public Values(string apiKey, string accessToken)
        {
            _apiKey = apiKey;
            _accessToken = accessToken;
        }

        /// <summary>
        /// Performs an operation.
        /// </summary>
        /// <returns>The status of the operation.</returns>
        public int Operation()
        {
            return _status;
        }

        /// <summary>
        /// Appends values to a range in a spreadsheet.
        /// </summary>
        /// <param name="spreadsheetId">The ID of the spreadsheet.</param>
        /// <param name="rng">The range to append values to.</param>
        /// <param name="json">The JSON representation of the values to append.</param>
        /// <param name="queryParameters">The query parameters.</param>
        /// <returns>The response from the API.</returns>
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

        /// <summary>
        /// Clears values from a spreadsheet.
        /// </summary>
        /// <param name="spreadsheetId">The ID of the spreadsheet.</param>
        /// <param name="json">The JSON representation of the clear request.</param>
        /// <param name="queryParameters">The query parameters.</param>
        /// <returns>The response from the API.</returns>
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

        /// <summary>
        /// Clears values from a spreadsheet based on a data filter.
        /// </summary>
        /// <param name="spreadsheetId">The ID of the spreadsheet.</param>
        /// <param name="json">The JSON representation of the clear request.</param>
        /// <param name="queryParameters">The query parameters.</param>
        /// <returns>The response from the API.</returns>
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

        /// <summary>
        /// Retrieves multiple ranges of values from a spreadsheet.
        /// </summary>
        /// <param name="spreadsheetId">The ID of the spreadsheet.</param>
        /// <param name="queryParameters">The query parameters.</param>
        /// <returns>The response from the API.</returns>
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

        /// <summary>
        /// Retrieves multiple ranges of values from a spreadsheet based on a data filter.
        /// </summary>
        /// <param name="spreadsheetId">The ID of the spreadsheet.</param>
        /// <param name="json">The JSON representation of the data filter request.</param>
        /// <param name="queryParameters">The query parameters.</param>
        /// <returns>The response from the API.</returns>
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

        /// <summary>
        /// Updates multiple ranges of values in a spreadsheet.
        /// </summary>
        /// <param name="spreadsheetId">The ID of the spreadsheet.</param>
        /// <param name="json">The JSON representation of the update request.</param>
        /// <param name="queryParameters">The query parameters.</param>
        /// <returns>The response from the API.</returns>
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

        /// <summary>
        /// Updates multiple ranges of values in a spreadsheet based on a data filter.
        /// </summary>
        /// <param name="spreadsheetId">The ID of the spreadsheet.</param>
        /// <param name="json">The JSON representation of the update request.</param>
        /// <param name="queryParameters">The query parameters.</param>
        /// <returns>The response from the API.</returns>
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

        /// <summary>
        /// Clears values from a range in a spreadsheet.
        /// </summary>
        /// <param name="spreadsheetId">The ID of the spreadsheet.</param>
        /// <param name="rng">The range to clear values from.</param>
        /// <param name="queryParameters">The query parameters.</param>
        /// <returns>The response from the API.</returns>
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

        /// <summary>
        /// Retrieves values from a range in a spreadsheet.
        /// </summary>
        /// <param name="spreadsheetId">The ID of the spreadsheet.</param>
        /// <param name="rng">The range to retrieve values from.</param>
        /// <param name="queryParameters">The query parameters.</param>
        /// <returns>The response from the API.</returns>
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

        /// <summary>
        /// Updates values in a range in a spreadsheet.
        /// </summary>
        /// <param name="spreadsheetId">The ID of the spreadsheet.</param>
        /// <param name="rng">The range to update values in.</param>
        /// <param name="json">The JSON representation of the values to update.</param>
        /// <param name="queryParameters">The query parameters.</param>
        /// <returns>The response from the API.</returns>
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
