using System;
using System.Runtime.InteropServices;

namespace GoogleApis.Oauth.Interfaces
{
    /// <summary>
    /// Interfaz para exponer los métodos y propiedades de FlowOauth a VBA.
    /// </summary>
    [ComVisible(true)]
    [Guid("e2b5f8a1-4b6e-4b8e-9c3e-2f8b5a1e4b6e")]
    public interface IFlowOauth
    {
        /// <summary>
        /// Obtiene el token de acceso.
        /// </summary>
        string AccessToken { get; }

        /// <summary>
        /// Obtiene la clave de API.
        /// </summary>
        string ApiKey { get; }

        /// <summary>
        /// Asigna el navegador.
        /// </summary>
        string Navigator { get; set; } 

        /// <summary>
        /// Inicializa el flujo de autenticación OAuth.
        /// </summary>
        /// <param name="credentialsClientPath">Ruta del archivo de credenciales del cliente.</param>
        /// <param name="credentialsTokenPath">Ruta del archivo de tokens de credenciales.</param>
        /// <param name="credentialsApiKeyPath">Ruta del archivo de clave de API.</param>
        /// <param name="scope">Ámbito de acceso solicitado.</param>
        void InitializeFlow(string credentialsClientPath, string credentialsTokenPath, string credentialsApiKeyPath, string scopes);

        /// <summary>
        /// Revoca el token de acceso.
        /// </summary>
        /// <param name="credentialsTokenPath">Ruta del archivo de tokens de credenciales.</param>
        /// <returns>True si el token se revocó correctamente, de lo contrario, false.</returns>
        bool RevokeToken();
    }
}
