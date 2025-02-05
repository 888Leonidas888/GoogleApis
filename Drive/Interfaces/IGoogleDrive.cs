using System;
using System.Runtime.InteropServices;
using GoogleApis.Drive.Component;

namespace GoogleApis.Drive.Interfaces
{
    /// <summary>
    /// Interfaz para interactuar con Google Drive.
    /// </summary>
    [ComVisible(true)]
    [Guid("09909001-1411-4a59-97e6-b6c937c22b05")]
    public interface IGoogleDrive
    {
        /// <summary>
        /// Obtiene la versión de la API de Google Drive.
        /// </summary>
        /// <returns>La versión de la API de Google Drive.</returns>
        string VersionApiGoogleDrive();

        /// <summary>
        /// Establece la conexión con el servicio de Google Drive.
        /// </summary>
        /// <param name="oFlowOauth">El objeto de flujo de autenticación OAuth.</param>
        void ConnectionService(object oFlowOauth);

        /// <summary>
        /// Obtiene la instancia de la clase Files para interactuar con los archivos de Google Drive.
        /// </summary>
        /// <returns>La instancia de la clase Files.</returns>
        Files Files();
    }
}
