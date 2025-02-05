using System;
using System.Runtime.InteropServices;

namespace GoogleApis.Sheet.Interfaces
{
    /// <summary>
    /// Interfaz para manipular los valores en una hoja de cálculo.
    /// </summary>
    [ComVisible(true)]
    [Guid("8c2e3e38-b3c5-4cf4-b7b0-5b963ef60349")]
    public interface IValues
    {
        /// <summary>
        /// Realiza una operación.
        /// </summary>
        /// <returns>El resultado de la operación.</returns>
        int Operation();

        /// <summary>
        /// Agrega valores a una hoja de cálculo.
        /// </summary>
        /// <param name="spreadsheetId">El ID de la hoja de cálculo.</param>
        /// <param name="rng">El rango de celdas donde se agregarán los valores.</param>
        /// <param name="json">Los valores en formato JSON.</param>
        /// <param name="queryParameters">Parámetros de consulta adicionales.</param>
        /// <returns>El resultado de la operación.</returns>
        string Append(string spreadsheetId, string rng, string json, Scripting.Dictionary queryParameters);

        /// <summary>
        /// Borra todos los valores de una hoja de cálculo.
        /// </summary>
        /// <param name="spreadsheetId">El ID de la hoja de cálculo.</param>
        /// <param name="json">Los valores en formato JSON.</param>
        /// <param name="queryParameters">Parámetros de consulta adicionales.</param>
        /// <returns>El resultado de la operación.</returns>
        string BatchClear(string spreadsheetId, string json, Scripting.Dictionary queryParameters);

        /// <summary>
        /// Borra los valores de una hoja de cálculo según un filtro de datos.
        /// </summary>
        /// <param name="spreadsheetId">El ID de la hoja de cálculo.</param>
        /// <param name="json">Los valores en formato JSON.</param>
        /// <param name="queryParameters">Parámetros de consulta adicionales.</param>
        /// <returns>El resultado de la operación.</returns>
        string BatchClearByDataFilter(string spreadsheetId, string json, Scripting.Dictionary queryParameters);

        /// <summary>
        /// Obtiene los valores de una hoja de cálculo.
        /// </summary>
        /// <param name="spreadsheetId">El ID de la hoja de cálculo.</param>
        /// <param name="queryParameters">Parámetros de consulta adicionales.</param>
        /// <returns>Los valores de la hoja de cálculo en formato JSON.</returns>
        string BatchGet(string spreadsheetId, Scripting.Dictionary queryParameters = null);

        /// <summary>
        /// Obtiene los valores de una hoja de cálculo según un filtro de datos.
        /// </summary>
        /// <param name="spreadsheetId">El ID de la hoja de cálculo.</param>
        /// <param name="json">Los valores en formato JSON.</param>
        /// <param name="queryParameters">Parámetros de consulta adicionales.</param>
        /// <returns>Los valores de la hoja de cálculo en formato JSON.</returns>
        string BatchGetByDataFilter(string spreadsheetId, string json, Scripting.Dictionary queryParameters = null);

        /// <summary>
        /// Actualiza los valores de una hoja de cálculo.
        /// </summary>
        /// <param name="spreadsheetId">El ID de la hoja de cálculo.</param>
        /// <param name="json">Los valores en formato JSON.</param>
        /// <param name="queryParameters">Parámetros de consulta adicionales.</param>
        /// <returns>El resultado de la operación.</returns>
        string BatchUpdate(string spreadsheetId, string json, Scripting.Dictionary queryParameters = null);

        /// <summary>
        /// Actualiza los valores de una hoja de cálculo según un filtro de datos.
        /// </summary>
        /// <param name="spreadsheetId">El ID de la hoja de cálculo.</param>
        /// <param name="json">Los valores en formato JSON.</param>
        /// <param name="queryParameters">Parámetros de consulta adicionales.</param>
        /// <returns>El resultado de la operación.</returns>
        string BatchUpdateByDataFilter(string spreadsheetId, string json, Scripting.Dictionary queryParameters = null);

        /// <summary>
        /// Borra los valores de una celda en una hoja de cálculo.
        /// </summary>
        /// <param name="spreadsheetId">El ID de la hoja de cálculo.</param>
        /// <param name="rng">El rango de celda a borrar.</param>
        /// <param name="queryParameters">Parámetros de consulta adicionales.</param>
        /// <returns>El resultado de la operación.</returns>
        string Clear(string spreadsheetId, string rng, Scripting.Dictionary queryParameters = null);

        /// <summary>
        /// Obtiene los valores de una celda en una hoja de cálculo.
        /// </summary>
        /// <param name="spreadsheetId">El ID de la hoja de cálculo.</param>
        /// <param name="rng">El rango de celda a obtener.</param>
        /// <param name="queryParameters">Parámetros de consulta adicionales.</param>
        /// <returns>El valor de la celda en formato JSON.</returns>
        string Get(string spreadsheetId, string rng, Scripting.Dictionary queryParameters = null);

        /// <summary>
        /// Actualiza el valor de una celda en una hoja de cálculo.
        /// </summary>
        /// <param name="spreadsheetId">El ID de la hoja de cálculo.</param>
        /// <param name="rng">El rango de celda a actualizar.</param>
        /// <param name="json">El nuevo valor en formato JSON.</param>
        /// <param name="queryParameters">Parámetros de consulta adicionales.</param>
        /// <returns>El resultado de la operación.</returns>
        string Update(string spreadsheetId, string rng, string json, Scripting.Dictionary queryParameters = null);
    }
}
