using Microsoft.Win32;
using System.Security.Principal;
using COMContract;

const string progId = ContractGuids.ProgId; //"MyCompany.MyExcelAddin";
const string clsid = $"{{{ContractGuids.Guid}}}";
const string friendlyName = ContractGuids.FriendlyName; //"My Excel Add-in";
const string comHostName = ContractGuids.ComHostName; // "MyExcelAddin.comhost.dll";
const string officeAddinsBase = @"SOFTWARE\Microsoft\Office\Excel\Addins";

var help = args.Contains("--help", StringComparer.OrdinalIgnoreCase);
if (help)
{
    Console.WriteLine("Usage: Registration.exe [--unregister] [--help]");
    Console.WriteLine("Options:");
    Console.WriteLine("  --unregister   Unregister the add-in.");
    Console.WriteLine("  --help         Show this help message.");
    return;
}

if (!IsAdministrator())
{
    Console.WriteLine("Please run this application as Administrator.");
    return;
}

var unregister = args.Contains("--unregister", StringComparer.OrdinalIgnoreCase);
try
{
    if (unregister)
    {
        UnregisterFromClassesRoot();
        UnregisterFromLocalMachine();
        Console.WriteLine("Add-in unregistered successfully.");
    }
    else
    {
        RegisterInClassesRoot();
        RegisterInLocalMachine();
        Console.WriteLine("Add-in registered successfully.");
    }
}
catch (FileNotFoundException)
{
        Console.WriteLine($"File not found. Make sure this app is in same directory as {comHostName} when executed.");
}
catch (Exception ex)
{
    Console.WriteLine($"Operation failed: {ex.Message}");
}

bool IsAdministrator() =>
    new WindowsPrincipal(WindowsIdentity.GetCurrent()!)
        .IsInRole(WindowsBuiltInRole.Administrator);

void RegisterInClassesRoot()
{
    var dllPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, comHostName);// @"C:\Path\To\MyExcelAddin.dll"; // Replace with actual path
    if (!File.Exists(dllPath)) throw new FileNotFoundException($"DLL not found: {dllPath}");

    using var progIdKey = Registry.ClassesRoot.CreateSubKey(progId);
    progIdKey?.SetValue(string.Empty, friendlyName);

    using var clsidSubKey = progIdKey?.CreateSubKey("CLSID");
    clsidSubKey?.SetValue(string.Empty, clsid);

    using var clsidKey = Registry.ClassesRoot.CreateSubKey($@"CLSID\{clsid}");
    clsidKey?.SetValue(string.Empty, friendlyName);

    using var inprocKey = clsidKey?.CreateSubKey("InprocServer32");
    inprocKey?.SetValue(string.Empty, dllPath); // Path to comhost.dll
    inprocKey?.SetValue("ThreadingModel", "Both");
    inprocKey?.SetValue("ProgID", progId);
}

void RegisterInLocalMachine()
{
    using var addinKey = Registry.LocalMachine.CreateSubKey($@"{officeAddinsBase}\{progId}");
    addinKey?.SetValue("Description", friendlyName);
    addinKey?.SetValue("FriendlyName", friendlyName);
    addinKey?.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
}

void UnregisterFromClassesRoot()
{
    Registry.ClassesRoot.DeleteSubKeyTree(progId, throwOnMissingSubKey: false);
    Registry.ClassesRoot.DeleteSubKeyTree($@"CLSID\{clsid}", throwOnMissingSubKey: false);
}

void UnregisterFromLocalMachine()
{
    Registry.LocalMachine.DeleteSubKeyTree($@"{officeAddinsBase}\{progId}", throwOnMissingSubKey: false);
}
