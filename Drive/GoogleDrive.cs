using System;
using System.Runtime.InteropServices;
using GoogleApis.Drive.Interfaces;
using GoogleApis.Drive.Component;
using GoogleApis.Drive.Common;
using GoogleApis.Oauth;

namespace GoogleApis
{    
    /// <summary>
    /// Clase que representa la interfaz de Google Drive.
    /// </summary>
    [ComVisible(true)]
    [Guid("9f8b1b7e-0c42-40e3-89a7-4bda8a8e2e30")]
    [ClassInterface(ClassInterfaceType.None)]
    public class GoogleDrive : GoogleDriveBase, IGoogleDrive
    {
        private const string _VERSION_API_GOOGLE_DRIVE = "v3";

        /// <summary>
        /// Obtiene la versión de la API de Google Drive.
        /// </summary>
        /// <returns>La versión de la API de Google Drive.</returns>
        public string VersionApiGoogleDrive()
        {
            return _VERSION_API_GOOGLE_DRIVE;
        }

        /// <summary>
        /// Establece la conexión con el servicio de Google Drive.
        /// </summary>
        /// <param name="oFlowOauth">El flujo de autenticación OAuth.</param>
        public void ConnectionService(FlowOauth oauth)
        {        
            _apiKey = oauth.ApiKey;
            _accessToken = oauth.AccessToken;
        }

        /// <summary>
        /// Obtiene una instancia de la clase Files para interactuar con los archivos de Google Drive.
        /// </summary>
        /// <returns>Una instancia de la clase Files.</returns>
        public Files Files()
        {
            return new Files(this._apiKey, this._accessToken);
        }
    }
}
