using System;
using System.Net.Http;
using System.Runtime.InteropServices;

namespace GoogleApis.Drive.Common
{
    /// <summary>
    /// Clase base para la interacción con la API de Google Drive.
    /// </summary>
    [ComVisible(false)]
    public class GoogleDriveBase
    {
        protected const string SERVICE_END_POINT = "https://www.googleapis.com/drive/v3/files";
        protected const string SERVICE_END_POINT_UPLOAD = "https://www.googleapis.com/upload/drive/v3/files";

        protected string _apiKey;
        protected string _accessToken;
        protected int _status;

        /// <summary>
        /// Crea los parámetros de consulta para la URL de la API de Google Drive.
        /// </summary>
        /// <param name="pathParameters">Parámetros de ruta opcionales.</param>
        /// <param name="queryParameters">Parámetros de consulta opcionales.</param>
        /// <param name="endPointUpload">Indica si se debe utilizar el punto final de carga.</param>
        /// <returns>La URL completa con los parámetros de consulta.</returns>
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

        /// <summary>
        /// Realiza una solicitud a la API de Google Drive.
        /// </summary>
        /// <param name="method">El método HTTP de la solicitud.</param>
        /// <param name="url">La URL de la solicitud.</param>
        /// <param name="body">El cuerpo de la solicitud.</param>
        /// <param name="headers">Los encabezados de la solicitud.</param>
        /// <returns>La respuesta de la solicitud.</returns>
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
