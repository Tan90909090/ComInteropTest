using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace ClientTestCSharpComImport;

[SupportedOSPlatform("windows")]

internal class TestComImport
{
    static void Main()
    {
        NativeMethods.CoCreateInstance(
            NativeMethods.CLSID_DummyNamespaceWalk,
            null,
            NativeMethods.CLSCTX_LOCAL_SERVER,
            typeof(INamespaceWalk).GUID,
            out object objNamespaceWalk);
        var pNamespaceWalk = (INamespaceWalk)objNamespaceWalk;

        NativeMethods.CoCreateInstance(
            NativeMethods.CLSID_DummyNamespaceWalkCB,
            null,
            NativeMethods.CLSCTX_LOCAL_SERVER,
            typeof(INamespaceWalkCB).GUID,
            out object objNamespaceWalkCB);
        var pNamespaceWalkCB = (INamespaceWalkCB)objNamespaceWalkCB;
        var pNamespaceWalkCB2 = (INamespaceWalkCB2)pNamespaceWalkCB;


        pNamespaceWalkCB.InitializeProgressDialog(out _, out _);
        pNamespaceWalkCB2.WalkComplete(0);
        pNamespaceWalk.Walk(null, 0, 0, pNamespaceWalkCB);

        _ = Marshal.ReleaseComObject(objNamespaceWalkCB);
        _ = Marshal.ReleaseComObject(objNamespaceWalk);

        // RCW解放後に使おうとすると「System.Runtime.InteropServices.InvalidComObjectException: 'COM object that has been separated from its underlying RCW cannot be used.'」
        // pNamespaceWalk.Walk(null, 0, 0, pNamespaceWalkCB);
    }
}

internal static class NativeMethods
{
#pragma warning disable SYSLIB1054 // Use 'LibraryImportAttribute' instead of 'DllImportAttribute' to generate P/Invoke marshalling code at compile time
    [DllImport("Ole32.dll", ExactSpelling = true, PreserveSig = false)]
    extern internal static void CoCreateInstance(
        in Guid rclsid,
        [MarshalAs(UnmanagedType.Interface)] object? pUnkOuter,
        uint dwClsContext,
        in Guid riid,
        [MarshalAs(UnmanagedType.Interface)] out object ppv);
#pragma warning restore SYSLIB1054 // Use 'LibraryImportAttribute' instead of 'DllImportAttribute' to generate P/Invoke marshalling code at compile time

    internal const uint CLSCTX_LOCAL_SERVER = 4;

    internal static readonly Guid CLSID_DummyNamespaceWalk = Guid.Parse("{3A9F4A2A-C7C6-4D9D-8F9E-D71EE470B57F}");
    internal static readonly Guid CLSID_DummyNamespaceWalkCB = Guid.Parse("{1B94A81D-A105-428D-8902-849F8D6483A2}");
}


[ComImport]
[Guid("57ced8a7-3f4a-432c-9350-30f24483f74f")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface INamespaceWalk
{
    void Walk(
        [MarshalAs(UnmanagedType.Interface)] object? punkToWalk,
        uint dwFlags,
        int cDepth,
        [MarshalAs(UnmanagedType.Interface)] INamespaceWalkCB pnswcb);
        
        void GetIDArrayResult(
            out uint pcItems,
            out IntPtr prgpidl); // actual is "PIDLIST_ABSOLUTE **"

};

#pragma warning disable SYSLIB1096 // Convert to 'GeneratedComInterface'
[ComImport]
[Guid("d92995f8-cf5e-4a76-bf59-ead39ea2b97e")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface INamespaceWalkCB
{
    void FoundItem(
        [MarshalAs(UnmanagedType.Interface)] object psf,
        IntPtr pidl);

    void EnterFolder(
        [MarshalAs(UnmanagedType.Interface)] object psf,
        IntPtr pidl);

    void LeaveFolder(
        [MarshalAs(UnmanagedType.Interface)] object psf,
        IntPtr pidl);

    // 引数は実際はLPWSTR *だけれど、outで受けるとBStrとして扱いそうなのでIntPtrでごまかす
    void InitializeProgressDialog(
        out IntPtr ppszTitle,
        out IntPtr ppszCancel);
}

// ComImport使用時は基底インターフェースの内容コピペが必要

[ComImport]
[Guid("7ac7492b-c38e-438a-87db-68737844ff70")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface INamespaceWalkCB2 : INamespaceWalkCB
{
    new void FoundItem(
        [MarshalAs(UnmanagedType.Interface)] object psf,
        IntPtr pidl);

    new void EnterFolder(
        [MarshalAs(UnmanagedType.Interface)] object psf,
        IntPtr pidl);

    new void LeaveFolder(
        [MarshalAs(UnmanagedType.Interface)] object psf,
        IntPtr pidl);

    // 引数は実際はLPWSTR *だけれど、outで受けるとBStrとして扱いそうなのでIntPtrでごまかす
    new void InitializeProgressDialog(
        out IntPtr ppszTitle,
        out IntPtr ppszCancel);

    void WalkComplete(int hr);
}
#pragma warning restore SYSLIB1096 // Convert to 'GeneratedComInterface'
