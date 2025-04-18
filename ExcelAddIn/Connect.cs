using COMContract;
using Extensibility;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace HcExcelAddIn;

[ComVisible(true)]
[Guid(ContractGuids.Guid)]
[ProgId(ContractGuids.ProgId)]
public class Connect : IDTExtensibility2, IConnect
{
    //[Guid(ContractGuids.AddInGuid)]
    private Application? _app;

    public void OnBeginShutdown(ref Array custom)
    {
        // This method is called when the add-in is being unloaded.
        // You can perform any necessary cleanup here.
    }

    public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
    {
        // This method is called when the add-in is loaded.
        // You can perform any initialization here.
        _app = application as Application;
    }

    public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
    {
        // This method is called when the add-in is unloaded.
        // You can perform any necessary cleanup here.
    }

    public void OnStartupComplete(ref Array custom)
    {
        // This method is called when the add-in has finished loading.
        // You can perform any final initialization here.

        if (_app == null)
        {
            throw new InvalidOperationException("Application object is not initialized.");
        }
        _app.ActiveSheet.Cells[1, 1].Value = "Hello, World!";
    }

    public void OnAddInsUpdate(ref Array custom)
    {
    }

    public void Test()
    {
        
    }
}