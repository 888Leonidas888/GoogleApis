using GoogleApis.Sheet.Interfaces;
using GoogleApis.Sheet.Components;
using GoogleApis.Sheet.Common;
using System;
using System.Runtime.InteropServices;

namespace GoogleApis.Sheet
{   
    [ComVisible(true)]
    [Guid("cdd4e85b-c33b-4e9b-b825-845307f9c190")]
    [ClassInterface(ClassInterfaceType.None)]
    public class GoogleSheets: GoogleSheetsBase, IGoogleSheets
    {    
        private const string _VERSION_API_GOOGLE_SHEETS = "v4";
        public string VersionApiGoogleSheets()
        {
            return _VERSION_API_GOOGLE_SHEETS;
        }
        public int Operation()
        {
            return _status;
        }
        public void ConnectionService(object oFlowOauth)
        {
            dynamic flow = oFlowOauth;

            _apiKey = flow.GetApiKey();
            _accessToken = flow.GetTokenAccess();
        }        
        public SpreadSheets SpreadSheets()
        {
            return new SpreadSheets(this._apiKey,this._accessToken);
        }
    }   
}
