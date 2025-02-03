using System;
using System.Runtime.InteropServices;

namespace GoogleApis.Sheet.Interfaces
{
    [ComVisible(true)]
    [Guid("89ccf222-d621-4736-8c53-8c347f7502b8")]
    public interface ISheets
    {
        int Operation();
        string CopyTo(string spreadsheetId, string sheetId, string json = null, Scripting.Dictionary queryParameters = null);
    }
}
