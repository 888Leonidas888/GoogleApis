using System;
using System.Runtime.InteropServices;
using GoogleApis.Drive.Component;

namespace GoogleApis.Drive.Interfaces
{
    [ComVisible(true)]
    [Guid("09909001-1411-4a59-97e6-b6c937c22b05")]
    public interface IGoogleDrive
    {
        string VersionApiGoogleDrive();
        void ConnectionService(object oFlowOauth);
        Files Files();
    }
}
