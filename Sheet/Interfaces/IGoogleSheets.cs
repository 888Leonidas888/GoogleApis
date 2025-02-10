using System;
using System.Runtime.InteropServices;
using GoogleApis.Sheet.Components;
using GoogleApis.Oauth;

namespace GoogleApis.Sheet.Interfaces
{
    /// <summary>
    /// Interfaz para interactuar con las hojas de cálculo de Google.
    /// </summary>
    [ComVisible(true)]
    [Guid("62e6123b-bd17-479d-9276-87536b16c584")]
    public interface IGoogleSheets
    {
        /// <summary>
        /// Obtiene la versión de la API de Google Sheets.
        /// </summary>
        /// <returns>La versión de la API de Google Sheets.</returns>
        string VersionApiGoogleSheets();

        /// <summary>
        /// Establece la conexión con el servicio de Google Sheets.
        /// </summary>
        /// <param name="oFlowOauth">El objeto de flujo de autenticación OAuth.</param>
        void ConnectionService(FlowOauth oauth);

        /// <summary>
        /// Obtiene una instancia de la clase SpreadSheets para interactuar con las hojas de cálculo.
        /// </summary>
        /// <returns>Una instancia de la clase SpreadSheets.</returns>
        SpreadSheets SpreadSheets();
    }
}
