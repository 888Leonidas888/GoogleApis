using System;
using System.Runtime.InteropServices;


namespace GoogleApis.Drive.Interfaces
{
    [ComVisible(true)]
    [Guid("13138978-4990-4b09-8e53-a226fee08e97")] // Un GUID único para la interfaz
    public interface IFiles
    {
        /// <summary>
        /// Realiza una operación genérica.
        /// </summary>
        /// <returns>Un entero que representa el resultado de la operación.</returns>
        int Operation();

        /// <summary>
        /// Conecta el servicio utilizando el flujo OAuth.
        /// </summary>
        /// <param name="oFlowOauth">El objeto de flujo OAuth.</param>
        void ConnectionService(object oFlowOauth);

        /// <summary>
        /// Copia un archivo en Google Drive.
        /// </summary>
        /// <param name="fileID">El ID del archivo a copiar.</param>
        /// <param name="fileObject">El objeto del archivo.</param>
        /// <param name="queryParameters">Parámetros de consulta opcionales.</param>
        /// <returns>El ID del nuevo archivo copiado.</returns>
        string Copy(string fileID, string fileObject, Scripting.Dictionary queryParameters = null);

        /// <summary>
        /// Elimina un archivo en Google Drive.
        /// </summary>
        /// <param name="fileID">El ID del archivo a eliminar.</param>
        /// <returns>Una cadena que representa el resultado de la operación.</returns>
        string Delete(string fileID);

        /// <summary>
        /// Vacía la papelera de Google Drive.
        /// </summary>
        /// <param name="queryParameters">Parámetros de consulta opcionales.</param>
        /// <returns>Una cadena que representa el resultado de la operación.</returns>
        string EmptyTrash(Scripting.Dictionary queryParameters = null);

        /// <summary>
        /// Exporta un archivo de Google Drive.
        /// </summary>
        /// <param name="fileID">El ID del archivo a exportar.</param>
        /// <param name="queryParameters">Parámetros de consulta.</param>
        /// <param name="pathFile">La ruta del archivo de destino.</param>
        /// <returns>Un valor booleano que indica si la exportación fue exitosa.</returns>
        bool Export(string fileID, Scripting.Dictionary queryParameters, string pathFile);

        /// <summary>
        /// Genera IDs para nuevos archivos en Google Drive.
        /// </summary>
        /// <param name="queryParameters">Parámetros de consulta opcionales.</param>
        /// <returns>Una cadena que representa los IDs generados.</returns>
        string GenerateIds(Scripting.Dictionary queryParameters = null);

        /// <summary>
        /// Obtiene los metadatos de un archivo en Google Drive.
        /// </summary>
        /// <param name="fileID">El ID del archivo.</param>
        /// <param name="queryParameters">Parámetros de consulta opcionales.</param>
        /// <returns>Una cadena que representa los metadatos del archivo.</returns>
        string GetMetadata(string fileID, Scripting.Dictionary queryParameters = null);

        /// <summary>
        /// Lista los archivos en Google Drive.
        /// </summary>
        /// <param name="queryParameters">Parámetros de consulta opcionales.</param>
        /// <returns>Una cadena que representa la lista de archivos.</returns>
        string List(Scripting.Dictionary queryParameters = null);

        /// <summary>
        /// Lista las etiquetas de un archivo en Google Drive.
        /// </summary>
        /// <param name="fileID">El ID del archivo.</param>
        /// <param name="queryParameters">Parámetros de consulta opcionales.</param>
        /// <returns>Una cadena que representa la lista de etiquetas.</returns>
        string ListLabels(string fileID, Scripting.Dictionary queryParameters = null);

        /// <summary>
        /// Actualiza un archivo en Google Drive.
        /// </summary>
        /// <param name="fileID">El ID del archivo.</param>
        /// <param name="fileObject">El objeto del archivo opcional.</param>
        /// <param name="queryParameters">Parámetros de consulta opcionales.</param>
        /// <returns>Una cadena que representa el resultado de la operación.</returns>
        string Update(string fileID, string fileObject = null, Scripting.Dictionary queryParameters = null);

        /// <summary>
        /// Descarga el contenido de un archivo en Google Drive.
        /// </summary>
        /// <param name="fileID">El ID del archivo.</param>
        /// <returns>Un valor booleano que indica si la descarga fue exitosa.</returns>
        bool DownLoadContentLink(string fileID);

        /// <summary>
        /// Descarga un archivo en Google Drive.
        /// </summary>
        /// <param name="fileID">El ID del archivo.</param>
        /// <param name="directory">El directorio de destino.</param>
        /// <returns>Un valor booleano que indica si la descarga fue exitosa.</returns>
        bool DownLoad(string fileID, string directory);

        /// <summary>
        /// Sube un archivo multimedia a Google Drive.
        /// </summary>
        /// <param name="pathFile">La ruta del archivo.</param>
        /// <param name="mimeType">El tipo MIME del archivo.</param>
        /// <returns>El ID del archivo subido.</returns>
        string UploadMedia(string pathFile = null, string mimeType = "application/octet-stream");

        /// <summary>
        /// Sube un archivo multipart a Google Drive.
        /// </summary>
        /// <param name="pathFile">La ruta del archivo.</param>
        /// <param name="fileObject">El objeto del archivo.</param>
        /// <returns>El ID del archivo subido.</returns>
        string UploadMultipart(string pathFile, string fileObject);
    }
}
