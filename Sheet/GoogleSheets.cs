using GoogleApis.Sheet.Interfaces;
using GoogleApis.Sheet.Components;
using GoogleApis.Sheet.Common;
using System;
using System.Runtime.InteropServices;

namespace GoogleApis.Sheet
{   
    /// <summary>
    /// Clase que representa la implementación de la interfaz IGoogleSheets.
    /// </summary>
    [ComVisible(true)]
    [Guid("cdd4e85b-c33b-4e9b-b825-845307f9c190")]
    [ClassInterface(ClassInterfaceType.None)]
    public class GoogleSheets : GoogleSheetsBase, IGoogleSheets
    {
        private const string _VERSION_API_GOOGLE_SHEETS = "v4";

        /// <summary>
        /// Obtiene la versión de la API de Google Sheets.
        /// </summary>
        /// <returns>La versión de la API de Google Sheets.</returns>
        public string VersionApiGoogleSheets()
        {
            return _VERSION_API_GOOGLE_SHEETS;
        }

        /// <summary>
        /// Establece la conexión con el servicio de Google Sheets.
        /// </summary>
        /// <param name="oFlowOauth">El objeto de flujo de autenticación OAuth.</param>
        public void ConnectionService(object oFlowOauth)
        {
            dynamic flow = oFlowOauth;

            _apiKey = flow.GetApiKey();
            _accessToken = flow.GetTokenAccess();
        }

        /// <summary>
        /// Obtiene una instancia de la clase SpreadSheets para interactuar con las hojas de cálculo.
        /// </summary>
        /// <returns>Una instancia de la clase SpreadSheets.</returns>
        public SpreadSheets SpreadSheets()
        {
            return new SpreadSheets(this._apiKey, this._accessToken);
        }
    }
}
