using Microsoft.Win32;
using COMContract;

namespace Register;

internal class Registration
{
    private static readonly string _addinPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"{ContractGuids.ComHostName}");
    private static readonly string _clsidKey = $@"CLSID\{{{ContractGuids.Guid}}}";

    // Make sure to not alter the order of the following 2 lines.
    private static readonly string _excelAddInRegistryPath = @"Software\Microsoft\Office\Excel\Addins";
    private static readonly string _addinKey = $@"{_excelAddInRegistryPath}\{ContractGuids.ProgId}";

    internal static void RegisterAddIn()
    {
        RegisterComClass();
        RegisterComClassStep2();
        RegisterExcelAddIn();
    }

    internal static void UnregisterAddIn()
    {
        UnregisterComClass();
        UnregisterComClassStep2();
        UnregisterExcelAddIn();
    }

    private static void RegisterComClass()
    {
        Console.WriteLine($"Registering COM class with CLSID: {ContractGuids.Guid}");   
        using RegistryKey key = Registry.ClassesRoot.CreateSubKey(_clsidKey);
        key.SetValue(null, ContractGuids.ProgId);

        Console.WriteLine($"Registering COM class with ProgId: {ContractGuids.ProgId}");
        using var inprocKey = key.CreateSubKey("InprocServer32");
        inprocKey.SetValue(null, _addinPath); // Path to comhost.dll
        inprocKey.SetValue("ThreadingModel", "Both");
        inprocKey.SetValue("ProgID", ContractGuids.ProgId);
    }

    private static void RegisterComClassStep2()
    {
        Console.WriteLine($"Registering COM class with ProgId: {ContractGuids.ProgId}");    
        using RegistryKey key = Registry.ClassesRoot.CreateSubKey(ContractGuids.ProgId);
        key.SetValue(null, ContractGuids.ProgId);

        Console.WriteLine($"Registering COM class with CLSID: {ContractGuids.Guid}");
        using var clsidKey = key.CreateSubKey("CLSID");
        clsidKey.SetValue(null, $"{{{ContractGuids.Guid}}}"); // Set the CLSID value
    }

    private static void RegisterExcelAddIn()
    {
        Console.WriteLine($"Registering Excel Add-In with ProgId: {ContractGuids.ProgId}");
        using var key = Registry.LocalMachine.CreateSubKey(_addinKey);
        key.SetValue("Description", ContractGuids.ProgId);
        key.SetValue("FriendlyName", ContractGuids.FriendlyName);
        key.SetValue("LoadBehavior", 3); // Load on startup
    }

    private static void UnregisterComClass()
    {
        Registry.ClassesRoot.DeleteSubKeyTree(_clsidKey, false);
    }

    private static void UnregisterComClassStep2()
    {
        Registry.ClassesRoot.DeleteSubKeyTree(ContractGuids.ProgId, false);
    }

    private static void UnregisterExcelAddIn()
    {
        Registry.LocalMachine.DeleteSubKeyTree(_addinKey, false);
    }
}