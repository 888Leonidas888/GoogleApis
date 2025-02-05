using System;
using System.Net.Http;
using System.Runtime.InteropServices;

namespace GoogleApis.Sheet.Common
{
    /// <summary>
    /// Clase base para las operaciones de Google Sheets.
    /// </summary>
    [ComVisible(false)]
    public class GoogleSheetsBase
    {
        protected const string SERVICE_END_POINT = "https://sheets.googleapis.com";
        protected string _apiKey;
        protected string _accessToken;
        protected int _status;

        /// <summary>
        /// Crea los parámetros de consulta para la URL.
        /// </summary>
        /// <param name="pathParameters">Parámetros de ruta opcionales.</param>
        /// <param name="queryParameters">Parámetros de consulta opcionales.</param>
        /// <returns>La URL completa con los parámetros de consulta.</returns>
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

        /// <summary>
        /// Realiza una solicitud HTTP.
        /// </summary>
        /// <param name="method">El método HTTP.</param>
        /// <param name="url">La URL de la solicitud.</param>
        /// <param name="body">El cuerpo de la solicitud.</param>
        /// <param name="headers">Los encabezados de la solicitud.</param>
        /// <returns>La respuesta HTTP.</returns>
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
