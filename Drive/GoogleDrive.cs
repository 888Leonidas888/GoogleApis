using System;
using System.Runtime.InteropServices;
using GoogleApis.Drive.Interfaces;
using GoogleApis.Drive.Component;
using GoogleApis.Drive.Common;

namespace GoogleApis
{    
    [ComVisible(true)]
    [Guid("9f8b1b7e-0c42-40e3-89a7-4bda8a8e2e30")]
    [ClassInterface(ClassInterfaceType.None)]
    public class GoogleDrive: GoogleDriveBase, IGoogleDrive
    {
        private const string _VERSION_API_GOOGLE_DRIVE = "v3";
        public string VersionApiGoogleDrive()
        {
            return _VERSION_API_GOOGLE_DRIVE;
        }       
        public void ConnectionService(object oFlowOauth)
        {
            dynamic flow = oFlowOauth;

            _apiKey = flow.GetApiKey();
            _accessToken = flow.GetTokenAccess();
        }        
        public Files Files()
        {
            return new Files(this._apiKey, this._accessToken);
        }
    }
}
