using System;
using System.Runtime.InteropServices;
using GoogleApis.Sheet.Components;

namespace GoogleApis.Sheet.Interfaces
{
    [ComVisible(true)]
    [Guid("6c92c2a1-5812-4ebc-9e72-b49431131c1c")]
    public interface ISpreadSheets
    {
        int Operation();
        string BatchUpdate(string spreadsheetId, string json, Scripting.Dictionary queryParameters = null);
        string Create(string json = null, Scripting.Dictionary queryParameters = null);
        string Get(string spreadsheetId, Scripting.Dictionary queryParameters = null);
        string GetByDataFilter(string spreadsheetId, string json, Scripting.Dictionary queryParameters = null);
        Sheets Sheets();
        Values Values();
    }
}
