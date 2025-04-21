using HcExcelAddIn.Controllers;
using Microsoft.Office.Core;
using System.Windows.Controls.Ribbon;

namespace HcExcelAddIn.Abstractions;

internal interface IRibbonController : IRibbonExtensibility
{
    void OnLoaded(IRibbonUI ribbonUI);
    void OnAction(IRibbonControl control);

}
