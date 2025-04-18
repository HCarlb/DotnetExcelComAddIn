using COMContract;
using Extensibility;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

// Project references required:
// Nuget Packages: stdole (from Microsoft)
// COM-reference: Interop.Microsoft.Office.Interop.Excel from Microsoft Excel 16.0 Object Library


namespace HcExcelAddIn;

[ComVisible(true)]
[Guid(ContractGuids.Guid)]
[ProgId(ContractGuids.ProgId)]
public class Connect : IDTExtensibility2
{
    private Application? _xlApp;

    public void OnBeginShutdown(ref Array custom)
    {
        // This method is called when the add-in is being unloaded.
        // You can perform any necessary cleanup here.
    }

    public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
    {
        // This method is called when the add-in is loaded.
        // You can perform any initialization here.
        _xlApp = application as Application;
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

        if (_xlApp == null) throw new InvalidOperationException("Application object is not initialized.");

        _xlApp.ActiveSheet.Cells[1, 1].Value = "Hello, World!";
    }

    public void OnAddInsUpdate(ref Array custom)
    {
        // This method is called when the add-in is updated.
        // You can perform any necessary updates here.
    }
}