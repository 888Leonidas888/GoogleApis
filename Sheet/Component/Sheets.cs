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
    [Guid("c7c37020-6701-4496-80c2-2f25b3e15ce8")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Sheets : GoogleSheetsBase, ISheets
    {
        /// <summary>
        /// Inicializa una nueva instancia de la clase Sheets.
        /// </summary>
        /// <param name="apiKey">La clave de API.</param>
        /// <param name="accessToken">El token de acceso.</param>
        public Sheets(string apiKey, string accessToken)
        {
            _apiKey = apiKey;
            _accessToken = accessToken;
        }

        /// <summary>
        /// Realiza una operación.
        /// </summary>
        /// <returns>El estado de la operación.</returns>
        public int Operation()
        {
            return _status;
        }

        /// <summary>
        /// Copia una hoja de cálculo a otra.
        /// </summary>
        /// <param name="spreadsheetId">El ID de la hoja de cálculo.</param>
        /// <param name="sheetId">El ID de la hoja.</param>
        /// <param name="json">El objeto JSON que contiene los datos de la copia.</param>
        /// <param name="queryParameters">Los parámetros de consulta adicionales.</param>
        /// <returns>El resultado de la copia.</returns>
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
