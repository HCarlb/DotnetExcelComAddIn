using System.Runtime.InteropServices;
using COMContract;

namespace HcExcelAddIn;

[ComVisible(true)]
[Guid(ContractGuids.Guid)]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal partial interface IConnect
{
    void Test();
}
