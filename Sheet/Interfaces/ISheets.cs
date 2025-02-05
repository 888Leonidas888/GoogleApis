using System;
using System.Runtime.InteropServices;

namespace GoogleApis.Sheet.Interfaces
{
    /// <summary>
    /// Interfaz para las hojas de cálculo.
    /// </summary>
    [ComVisible(true)]
    [Guid("89ccf222-d621-4736-8c53-8c347f7502b8")]
    public interface ISheets
    {
        /// <summary>
        /// Realiza una operación.
        /// </summary>
        /// <returns>El resultado de la operación.</returns>
        int Operation();

        /// <summary>
        /// Copia una hoja de cálculo a otra.
        /// </summary>
        /// <param name="spreadsheetId">El ID de la hoja de cálculo a copiar.</param>
        /// <param name="sheetId">El ID de la hoja a copiar.</param>
        /// <param name="json">El JSON de configuración opcional.</param>
        /// <param name="queryParameters">Los parámetros de consulta opcionales.</param>
        /// <returns>El resultado de la copia.</returns>
        string CopyTo(string spreadsheetId, string sheetId, string json = null, Scripting.Dictionary queryParameters = null);
    }
}
