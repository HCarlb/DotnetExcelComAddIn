using Microsoft.Office.Core;
using System.Reflection;
using XlApplication = Microsoft.Office.Interop.Excel.Application;

namespace HcExcelAddIn;

internal sealed class RibbonController(XlApplication xlApp, string ribbonName) : IRibbonExtensibility, IDisposable
{
    private readonly XlApplication _xlApp = xlApp;
    private readonly string _ribbonName = ribbonName;
    private readonly string _ribbonPath = "Ribbons";

    private IRibbonUI? _ribbon;
    private bool _disposed = false;
    
    public void OnLoaded(IRibbonUI ribbon)
    {
        _ribbon = ribbon;
        Log.Information("Ribbon {0} loaded.", _ribbonName);
    }

    public string GetCustomUI(string RibbonID)
    {
        var ribbonFullName = GetRibbonResourceName(_ribbonName, _ribbonPath);
        return GetEmbeddedResource(ribbonFullName);
    }

    public void OnAction(IRibbonControl control)
    {
        // Handle ribbon button click events here
        switch (control.Id)
        {
            case "button1":
                //MessageBox.Show("Button 1 clicked!", "Ribbon Button Clicked", MessageBoxButtons.OK, MessageBoxIcon.Information);
                var x = new MainView();
                x.Show();
                break;
            default:
                Log.Warning("Unknown control ID: {0}", control.Id);
                break;
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        
        // Cleanup future resources or event handlers
        _ribbon = null;
        _disposed = true;
        Log.Debug("RibbonController for {0} disposed.", _ribbonName);
    }

    /// <summary>
    /// Gets the full name of the embedded resource for the ribbon.
    /// Ex. AppName.Folder.Ribbon.xml
    /// </summary>
    /// <param name="name"></param>
    /// <param name="resourcePath"></param>
    /// <returns></returns>
    private static string GetRibbonResourceName(string name, string resourcePath)
    {
        var assembly = Assembly.GetExecutingAssembly();
        var resourceNames = assembly.GetManifestResourceNames();
        return resourceNames.Single(str => str.EndsWith(name) && str.Contains(resourcePath));
    }

    /// <summary>
    /// Gets the embedded resource as a string.
    /// </summary>
    /// <param name="resourceName"></param>
    /// <returns></returns>
    private static string GetEmbeddedResource(string resourceName)
    {
        Log.Debug("Loading embedded resource: {0}", resourceName);
        using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName);
        if (stream == null)
        {
            Log.Error("Resource '{0}' not found.", resourceName);
            return string.Empty;
        }

        Log.Debug("Resource '{0}' found.", resourceName);
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }
}
