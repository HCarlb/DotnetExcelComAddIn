using COMContract;
using Extensibility;
//using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Runtime.InteropServices;
using stdole;
using Serilog;

// Project references required:
// Nuget Packages: stdole (from Microsoft)
// COM-reference: Interop.Microsoft.Office.Interop.Excel from Microsoft Excel 16.0 Object Library


namespace HcExcelAddIn;

[ComVisible(true)]
[Guid(ContractGuids.Guid)]
[ProgId(ContractGuids.ProgId)]
public class Connect : IDTExtensibility2 //, IRibbonExtensibility //, ICustomTaskPaneConsumer
{
    private Application? _xlApp;

    /* 
     * ################################################################################################################################
     * 
     * Connect() 
     * { 
     *      // Warning!!!! The constructor seem to make the addon not load in Excel.
     *      // Do not use!!!!!
     * }  
     * ################################################################################################################################
    */

    private static void ConfigureLogger()
    {
        var folder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? $"{ContractGuids.ProgId}_logs";
        var path = Path.Combine(folder, "log-.txt" );

        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Debug()
            .WriteTo.File(path, rollingInterval: RollingInterval.Day)
            .CreateLogger();

        Log.Information("Logger configured.");
    }

    public void OnBeginShutdown(ref Array custom)
    {
        //Log.Information("Add-in is being unloaded.");
        // This method is called when the add-in is being unloaded.
        // You can perform any necessary cleanup here.

        Log.CloseAndFlush();
    }

    public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
    {
        // This method is called when the add-in is loaded.
        // You can perform any initialization here.
        Log.Information("Add-in is being loaded.");

        ConfigureLogger();

        _xlApp = application as Application;
    }

    public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
    {
        // This method is called when the add-in is unloaded.
        // You can perform any necessary cleanup here.
        Log.Information("Add-in is being unloaded.");
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

        Log.Information("Greeting the user with Hello World!");
        _xlApp.ActiveSheet.Cells[1, 1].Value = "Hello, World!";


        Log.Information("Add-in has finished loading.");
    }

    public void OnAddInsUpdate(ref Array custom)
    {
        // This method is called when the add-in is updated.
        // You can perform any necessary updates here.
        Log.Information("Add-in has been updated.");
    }

     // Utility to retrieve the embedded resource as a string
    //private static string GetEmbeddedResource(string resourceName)
    //{
    //    //Log.Information($"Loading embedded resource: {resourceName}");
    //    using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName) ?? throw new InvalidOperationException($"Resource '{resourceName}' not found.");
    //    using var reader = new StreamReader(stream);
    //    return reader.ReadToEnd();
    //}

    //public string GetCustomUI(string RibbonID)
    //{
    //    return GetEmbeddedResource("Ribbon.xml");
    //}

    //public void CTPFactoryAvailable(ICTPFactory CTPFactoryInst)
    //{
    //   // SidePanelManager.Initialize(CTPFactoryInst);
    //}
}