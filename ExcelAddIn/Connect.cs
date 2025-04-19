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
    private readonly string _ribbonName = "Ribbon.xml";
    private readonly string _ribbonPath = "Ribbons";
    private XlApplication? _xlApp;

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
        // This method is called when the add-in is being unloaded.
        // You can perform any necessary cleanup here.

        Log.CloseAndFlush();
    }

    public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
    {
        // Setup Logging
        Log.Logger = Configuration.ConfigureLogger();
        Log.Information("Add-in is being loaded.");

        // This method is called when the add-in is loaded.
        // You can perform any initialization here.

        _xlApp = application as XlApplication;
    }

    public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
    {
        // This method is called when the add-in is unloaded.
        // You can perform any necessary cleanup here.
        Log.Debug("Add-in is being unloaded.");
    }

    public void OnStartupComplete(ref Array custom)
    {
        // This method is called when the add-in has finished loading.
        // You can perform any final initialization here.
        if (_xlApp == null)
        {
            Log.Error("Application object is not initialized.");
            return;
        }

        Log.Debug("Greeting the user with Hello World!");
        _xlApp.ActiveSheet.Cells[1, 1].Value = "Hello, World!";
        
        Log.Information("Add-in has finished loading.");
    }

    public void OnAddInsUpdate(ref Array custom)
    {
        // This method is called when the add-in is updated.
        // You can perform any necessary updates here.
        Log.Debug("Add-in has been updated.");
    }

    public string GetCustomUI(string RibbonID)
    {
        try
        {
            return RibbonExtensions.GetRibbonXML(_ribbonName, _ribbonPath);
        }
        catch (Exception ex)
        {
            Log.Error("Error loading ribbon XML: {0}", ex.Message);
            return string.Empty;
        }
    }

    public void CTPFactoryAvailable(ICTPFactory CTPFactoryInst)
    {
        // This method is called when the CTP factory is available.
        // CPT factory is used to create custom task panes.
        Log.Debug("CTPFactory available.");
    }

    public void OnRibbonLoaded(IRibbonUI ribbonUI)
    {
        // This method is called when the ribbon is loaded from onLoad="OnRibbonLoad" in the xml ribbon.
        // You can perform any necessary initialization here.
        Log.Debug("Ribbon loaded.");

        _ = ribbonUI;   // To discard the warning about unused variable.
    }
    
    public void OnButtonClick(IRibbonControl control)
    {   
        // This method is called when the button is clicked.
        // You can perform any necessary actions here.
        Log.Debug("Button clicked: {0}", control.Id);
        if (_xlApp is null)
        {
            Log.Error("Application object is not initialized.");
            return;
        }

        _xlApp.ActiveSheet.Cells[1, 2].Value = $"Button [{control.Id}] was Clicked";

        MessageBox.Show($"Button [{control.Id}] was Clicked", "Button Clicked", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
}