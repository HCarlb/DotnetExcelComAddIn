namespace COMContract;

/// <summary>
/// This class contains the GUIDs and other constants used in the COM Add-in and app to Register the addin.
/// It is used to register the add-in with the Windows registry and to identify the add-in in Excel.
/// </summary>
public static class ContractGuids
{
    // This is the GUID for the COM Add-in. It must be unique and should not be changed.
    // AppName.ClassName
    public const string ProgId = "HcExcelAddIn.Connect";

    // Name of the comhost.dll file
    // AppName
    public const string ComHostName = "HcExcelAddIn.comhost.dll";

    // Name of the add-in. This is the name that will be displayed in the Excel Add-ins dialog.
    // Can be whatever you want.
    public const string FriendlyName = "HcExcel Addin. This will be visible in Excel.";

    // This is the GUID for the COM Add-in. It must be unique and should not be changed.
    // Must be re-generated if you use this project as a template for new projects, to avoid overwrites in windows registry.
    public const string Guid = "75CDAAAE-EF6C-4433-A059-F1FA2ED46F47";
}
