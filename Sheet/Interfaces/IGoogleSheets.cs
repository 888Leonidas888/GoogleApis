using System;
using System.Runtime.InteropServices;
using GoogleApis.Sheet.Components;

namespace GoogleApis.Sheet.Interfaces
{
    [ComVisible(true)]
    [Guid("62e6123b-bd17-479d-9276-87536b16c584")]
    public interface IGoogleSheets
    {
        string VersionApiGoogleSheets();
        void ConnectionService(object oFlowOauth);
        SpreadSheets SpreadSheets();
    }
}
