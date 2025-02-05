using System;
using System.Runtime.InteropServices;
using GoogleApis.Sheet.Components;

namespace GoogleApis.Sheet.Interfaces
{
    /// <summary>
    /// Interfaz para interactuar con hojas de cálculo de Google.
    /// </summary>
    [ComVisible(true)]
    [Guid("6c92c2a1-5812-4ebc-9e72-b49431131c1c")]
    public interface ISpreadSheets
    {
        /// <summary>
        /// Realiza una operación.
        /// </summary>
        /// <returns>El resultado de la operación.</returns>
        int Operation();

        /// <summary>
        /// Actualiza en lote una hoja de cálculo.
        /// </summary>
        /// <param name="spreadsheetId">El ID de la hoja de cálculo.</param>
        /// <param name="json">El JSON con los datos a actualizar.</param>
        /// <param name="queryParameters">Los parámetros de consulta opcionales.</param>
        /// <returns>El resultado de la actualización en lote.</returns>
        string BatchUpdate(string spreadsheetId, string json, Scripting.Dictionary queryParameters = null);

        /// <summary>
        /// Crea una nueva hoja de cálculo.
        /// </summary>
        /// <param name="json">El JSON con los datos de la nueva hoja de cálculo.</param>
        /// <param name="queryParameters">Los parámetros de consulta opcionales.</param>
        /// <returns>El ID de la nueva hoja de cálculo.</returns>
        string Create(string json = null, Scripting.Dictionary queryParameters = null);

        /// <summary>
        /// Obtiene una hoja de cálculo existente.
        /// </summary>
        /// <param name="spreadsheetId">El ID de la hoja de cálculo.</param>
        /// <param name="queryParameters">Los parámetros de consulta opcionales.</param>
        /// <returns>El JSON con los datos de la hoja de cálculo.</returns>
        string Get(string spreadsheetId, Scripting.Dictionary queryParameters = null);

        /// <summary>
        /// Obtiene una hoja de cálculo existente filtrada por datos.
        /// </summary>
        /// <param name="spreadsheetId">El ID de la hoja de cálculo.</param>
        /// <param name="json">El JSON con los datos de filtrado.</param>
        /// <param name="queryParameters">Los parámetros de consulta opcionales.</param>
        /// <returns>El JSON con los datos de la hoja de cálculo filtrada.</returns>
        string GetByDataFilter(string spreadsheetId, string json, Scripting.Dictionary queryParameters = null);

        /// <summary>
        /// Obtiene las hojas de cálculo.
        /// </summary>
        /// <returns>El objeto Sheets para interactuar con las hojas de cálculo.</returns>
        Sheets Sheets();

        /// <summary>
        /// Obtiene los valores de las celdas de una hoja de cálculo.
        /// </summary>
        /// <returns>El objeto Values para interactuar con los valores de las celdas.</returns>
        Values Values();
    }
}
