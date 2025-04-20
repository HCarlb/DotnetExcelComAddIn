using Extensibility;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using XlApplication = Microsoft.Office.Interop.Excel.Application;

// Project references required:
// Nuget Packages: stdole (from Microsoft)
// COM-reference: Interop.Microsoft.Office.Interop.Excel from Microsoft Excel 16.0 Object Library
// COM-reference: Microsoft.Office.Core from Microsoft Office 16.0 Object Library


namespace HcExcelAddIn;

[ComVisible(true)]
[Guid(ContractGuids.Guid)]
[ProgId(ContractGuids.ProgId)]
public class Connect : IDTExtensibility2 , IRibbonExtensibility, ICustomTaskPaneConsumer
{
    private XlApplication? _xlApp;
    private RibbonController? _ribbonController;
    /* 
     * ################################################################################################################################
     * 
     * Connect() 
     * { 
     *      // Warning!!!! The constructor seem to make the addon not load in Excel.
     *      // Do not use(for now)!!!!! 
     * }  
     * ################################################################################################################################
    */

    public void OnBeginShutdown(ref Array custom)
    {
        Log.Information("Add-in is being unloaded.");
        Log.CloseAndFlush();
    }

    public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
    {
        // Setup Logging
        Log.Logger = Configuration.ConfigureLogger();
        Log.Information("Add-in is being loaded.");

        // Initialize the excel application object
        _xlApp = application as XlApplication;

        if (_xlApp == null)
        {
            Log.Error("Application object is not initialized.");
            return;
        }

        // RibbonController is used to manage the custom ribbon UI
        _ribbonController = new RibbonController(_xlApp, "Ribbon.xml");
    }

    public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
    {
        Log.Debug("Add-in is being unloaded.");

        // Clean up
        if (_ribbonController is IDisposable d) d.Dispose();
        _ribbonController = null;
        _xlApp = null;
    }

    public void OnStartupComplete(ref Array custom)
    {
        Log.Debug("Greeting the user with Hello World!");
        _xlApp!.ActiveSheet.Cells[1, 1].Value = "Hello, World!";
        
        Log.Information("Add-in has finished loading.");
    }

    public void OnAddInsUpdate(ref Array custom)
    {
        Log.Debug("Add-in has been updated.");
    }

    public void CTPFactoryAvailable(ICTPFactory CTPFactoryInst)
    {
        Log.Debug("CTPFactory available.");
    }

    #region RibbonController
    public string GetCustomUI(string RibbonID) => _ribbonController?.GetCustomUI(RibbonID) ?? string.Empty;
    public void OnRibbonLoaded(IRibbonUI ribbonUI) => _ribbonController?.OnLoaded(ribbonUI);
    public void OnAction(IRibbonControl control) => _ribbonController?.OnAction(control);
    #endregion RibbonController
}