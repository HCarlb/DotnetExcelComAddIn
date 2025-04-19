using Microsoft.Office.Core;
using XlApplication = Microsoft.Office.Interop.Excel.Application;

namespace HcExcelAddIn;

public sealed class RibbonController(XlApplication xlApp) : IRibbonExtensibility, IDisposable
{
    private readonly string _ribbonName = "Ribbon.xml";
    private readonly string _ribbonPath = "Ribbons";

    private readonly XlApplication _xlApp = xlApp;
    private IRibbonUI? _ribbon;
    private bool _disposed = false;
    private bool _loaded = false;

    public void OnLoaded(IRibbonUI ribbon)
    {
        _ribbon = ribbon;
        _loaded = true;
        Log.Information("Ribbon loaded.");
    }

    public string GetCustomUI(string RibbonID) => RibbonExtensions.GetRibbonXML(_ribbonName, _ribbonPath);

    public void OnAction(IRibbonControl control)
    {
        // Handle ribbon button click events here
        switch (control.Id)
        {
            case "button1":
                MessageBox.Show("Button 1 clicked!", "Ribbon Button Clicked", MessageBoxButtons.OK, MessageBoxIcon.Information);
                break;
            default:
                Log.Warning($"Unknown control ID: {control.Id}");
                break;
        }
    }

    public void Dispose()
    {
        if (!_disposed)
        {
            // Cleanup future resources or event handlers
            _loaded = false;
            _ribbon = null;
            _disposed = true;
            Log.Debug("RibbonController disposed.");
        }
    }
}
