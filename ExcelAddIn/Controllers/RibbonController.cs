using Microsoft.Office.Core;
using System.Reflection;

namespace HcExcelAddIn.Controllers;
internal sealed class RibbonController(ExcelApplication excelApplication, ILogger logger, IServiceProvider serviceProvider) : IRibbonController
{
    private readonly string _ribbonName = "Ribbon.xml";

    private readonly ExcelApplication _xlApp = excelApplication;
    private readonly ILogger _logger = logger;
    private readonly IServiceProvider _serviceProvider = serviceProvider;
    private IRibbonUI? _ribbon;

    public void OnLoaded(IRibbonUI ribbon)
    {
        _ribbon = ribbon;
    }

    public string GetCustomUI(string RibbonID) => Assembly.GetExecutingAssembly().GetEmbeddedResource(_ribbonName);

    public void OnAction(IRibbonControl control)
    {
        // Handle ribbon button click events here
        switch (control.Id)
        {
            case "button1":
                // WinForms MessageBox example
                MessageBox.Show("Button 1 clicked!", "Ribbon Button Clicked", MessageBoxButtons.OK, MessageBoxIcon.Information);
                break;
            case "button2":
                // WPF Window example
                _serviceProvider.GetService<MainView>()?.ShowDialog();
                break;
            default:
                _logger.Warning("Unknown control ID: {0}", control.Id);
                break;
        }
    }
}
