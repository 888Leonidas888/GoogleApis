using System;
using System.Runtime.InteropServices;
using GoogleApis.Sheet.Interfaces;
using GoogleApis.Sheet.Common;

namespace GoogleApis.Sheet.Components
{
    /// <summary>
    /// Clase que representa las hojas de cálculo de Google.
    /// </summary>
    [ComVisible(true)]
    [Guid("38ffbe19-2ef4-484d-9c8b-e742e56ec5ee")]
    [ClassInterface(ClassInterfaceType.None)]
    public class SpreadSheets : GoogleSheetsBase, ISpreadSheets
    {
        /// <summary>
        /// Inicializa una nueva instancia de la clase SpreadSheets.
        /// </summary>
        /// <param name="apiKey">La clave de API.</param>
        /// <param name="accessToken">El token de acceso.</param>
        public SpreadSheets(string apiKey, string accessToken)
        {
            _apiKey = apiKey;
            _accessToken = accessToken;
        }

        /// <summary>
        /// Realiza una operación y devuelve el estado.
        /// </summary>
        /// <returns>El estado de la operación.</returns>
        public int Operation()
        {
            return _status;
        }

        /// <summary>
        /// Realiza una actualización por lotes en una hoja de cálculo.
        /// </summary>
        /// <param name="spreadsheetId">El ID de la hoja de cálculo.</param>
        /// <param name="json">El JSON de actualización.</param>
        /// <param name="queryParameters">Los parámetros de consulta opcionales.</param>
        /// <returns>La respuesta de la actualización por lotes.</returns>
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

        /// <summary>
        /// Crea una nueva hoja de cálculo.
        /// </summary>
        /// <param name="json">El JSON de creación.</param>
        /// <param name="queryParameters">Los parámetros de consulta opcionales.</param>
        /// <returns>La respuesta de la creación de la hoja de cálculo.</returns>
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

        /// <summary>
        /// Obtiene una hoja de cálculo existente.
        /// </summary>
        /// <param name="spreadsheetId">El ID de la hoja de cálculo.</param>
        /// <param name="queryParameters">Los parámetros de consulta opcionales.</param>
        /// <returns>La respuesta de la obtención de la hoja de cálculo.</returns>
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

        /// <summary>
        /// Obtiene una hoja de cálculo existente filtrada por datos.
        /// </summary>
        /// <param name="spreadsheetId">El ID de la hoja de cálculo.</param>
        /// <param name="json">El JSON de filtrado de datos.</param>
        /// <param name="queryParameters">Los parámetros de consulta opcionales.</param>
        /// <returns>La respuesta de la obtención de la hoja de cálculo filtrada por datos.</returns>
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

        /// <summary>
        /// Obtiene una instancia de la clase Sheets.
        /// </summary>
        /// <returns>Una instancia de la clase Sheets.</returns>
        public Sheets Sheets()
        {
            return new Sheets(this._apiKey, this._accessToken);
        }

        /// <summary>
        /// Obtiene una instancia de la clase Values.
        /// </summary>
        /// <returns>Una instancia de la clase Values.</returns>
        public Values Values()
        {
            return new Values(this._apiKey, this._accessToken);
        }
    }
}
