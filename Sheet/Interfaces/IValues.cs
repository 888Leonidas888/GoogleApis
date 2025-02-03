using System;
using System.Runtime.InteropServices;

namespace GoogleApis.Sheet.Interfaces
{
    [ComVisible(true)]
    [Guid("8c2e3e38-b3c5-4cf4-b7b0-5b963ef60349")]
    public interface IValues
    {
        int Operation();
        string Append(string spreadsheetId, string rng, string json, Scripting.Dictionary queryParameters);
        string BatchClear(string spreadsheetId, string json, Scripting.Dictionary queryParameters);
        string BatchClearByDataFilter(string spreadsheetId, string json, Scripting.Dictionary queryParameters);
        string BatchGet(string spreadsheetId, Scripting.Dictionary queryParameters = null);
        string BatchGetByDataFilter(string spreadsheetId, string json, Scripting.Dictionary queryParameters = null);
        string BatchUpdate(string spreadsheetId, string json, Scripting.Dictionary queryParameters = null);
        string BatchUpdateByDataFilter(string spreadsheetId, string json, Scripting.Dictionary queryParameters = null);
        string Clear(string spreadsheetId, string rng, Scripting.Dictionary queryParameters = null);
        string Get(string spreadsheetId, string rng, Scripting.Dictionary queryParameters = null);
        string Update(string spreadsheetId, string rng, string json, Scripting.Dictionary queryParameters = null);
    }
}
