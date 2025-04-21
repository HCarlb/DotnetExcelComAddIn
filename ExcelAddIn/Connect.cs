using Extensibility;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;

// Project references required:
// Nuget Packages: stdole (from Microsoft)
// COM-reference: Interop.Microsoft.Office.Interop.Excel from Microsoft Excel 16.0 Object Library
// COM-reference: Microsoft.Office.Core from Microsoft Office 16.0 Object Library

namespace HcExcelAddIn;

[ComVisible(true)]
[Guid(ContractGuids.Guid)]
[ProgId(ContractGuids.ProgId)]
public sealed class Connect : IDTExtensibility2 , IRibbonExtensibility, ICustomTaskPaneConsumer
{
    private ExcelApplication? _xlApp;
    private IRibbonController? _ribbonController;
    private IServiceProvider? _serviceProvider;
    private ILogger? _logger;

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

    public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
    {
        // Initialize the excel application object
        _xlApp = application as ExcelApplication;
        if (_xlApp == null) return;

        _serviceProvider = Configuration.ConfigureServices(_xlApp);
        _logger = _serviceProvider.GetService<ILogger>();
        _ribbonController = _serviceProvider.GetService<IRibbonController>();
    }

    public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
    {
        // Clean up
        _ribbonController = null;
        _xlApp = null;
    }

    public void OnStartupComplete(ref Array custom)
    {
        if (_xlApp == null) return;
        _xlApp.ActiveSheet.Cells[1, 1].Value = "Hello, World!";
    }

    public void OnAddInsUpdate(ref Array custom) {}
    public void OnBeginShutdown(ref Array custom) { }
    public void CTPFactoryAvailable(ICTPFactory CTPFactoryInst) { }

    #region RibbonController
    public string GetCustomUI(string RibbonID) => _ribbonController?.GetCustomUI(RibbonID) ?? string.Empty;
    public void OnRibbonLoaded(IRibbonUI ribbonUI) => _ribbonController?.OnLoaded(ribbonUI);
    public void OnAction(IRibbonControl control) => _ribbonController?.OnAction(control);

    #endregion RibbonController
}