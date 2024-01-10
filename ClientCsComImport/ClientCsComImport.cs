using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace ClientTestCSharpComImport;

[SupportedOSPlatform("windows")]

internal class ClientCsComImport
{
    [STAThread]
    static void Main()
    {
        UseCom();
        GC.Collect();
        GC.WaitForPendingFinalizers(); // Marshal.ReleaseComObject未実施でもここでRCWのFinalizeが呼び出されて各種ネイティブCOMオブジェクトはReleaseされます
        GC.Collect();
    }

    private static void UseCom()
    {
        // なお、ComImportAttributeの場合は、CLSCTX_INPROC_SERVERでの生成はCoCreateInstanceをP/Invokeする以外の方法もあります
        if (false)
        {
#pragma warning disable CS0162 // Unreachable code detected
            object o1 = Activator.CreateInstance(Type.GetTypeFromCLSID(NativeMethods.CLSID_DummyNamespaceWalk)!)!;
            object o2 = new DummyNamespaceWalk(); // ComImportAttributeがついたclass
            object o3 = new INamespaceWalk(); // CoClassAttributeがついたinterface、もちろん普通の実装ではCoClass属性なんてありません
#pragma warning restore CS0162 // Unreachable code detected
        }

        // COMオブジェクト生成
        NativeMethods.CoCreateInstance(
            NativeMethods.CLSID_DummyNamespaceWalk,
            null,
            NativeMethods.CLSCTX_INPROC_SERVER,
            typeof(INamespaceWalk).GUID,
            out object objNamespaceWalk);
        var pNamespaceWalk = (INamespaceWalk)objNamespaceWalk;

        NativeMethods.CoCreateInstance(
            NativeMethods.CLSID_DummyNamespaceWalkCB,
            null,
            NativeMethods.CLSCTX_INPROC_SERVER,
            typeof(INamespaceWalkCB).GUID,
            out object objNamespaceWalkCB);
        var pNamespaceWalkCB = (INamespaceWalkCB)objNamespaceWalkCB;
        var pNamespaceWalkCB2 = (INamespaceWalkCB2)pNamespaceWalkCB;

        // 適当に使う
        pNamespaceWalkCB.InitializeProgressDialog(out _, out _);
        pNamespaceWalkCB2.WalkComplete(0);
        pNamespaceWalk.Walk(null, 0, 0, pNamespaceWalkCB);

        // 使い終わったので解放
        Console.Write("""
            Do you test to what hapen if you explicitly call Marshal.ReleaseComObject?
            0: Don't call Marshal.ReleaseComObject. RCW's Finalizer will be called later.
            1: Call Marshal.ReleaseComObject.
            > 
            """);
        _ = int.TryParse(Console.ReadLine(), out int choise);
        if(choise == 1)
        {
            _ = Marshal.ReleaseComObject(objNamespaceWalkCB);
            _ = Marshal.ReleaseComObject(objNamespaceWalk);
        }

        // Marshal.ReleaseComObject後にRCWを使おうとすると、適切に「System.Runtime.InteropServices.InvalidComObjectException: 'COM object that has been separated from its underlying RCW cannot be used.'」例外が発生します。安全です。
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

    internal const uint CLSCTX_INPROC_SERVER = 1;

    internal static readonly Guid CLSID_DummyNamespaceWalk = Guid.Parse("{3A9F4A2A-C7C6-4D9D-8F9E-D71EE470B57F}");
    internal static readonly Guid CLSID_DummyNamespaceWalkCB = Guid.Parse("{1B94A81D-A105-428D-8902-849F8D6483A2}");
}


[ComImport]
[Guid("57ced8a7-3f4a-432c-9350-30f24483f74f")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[CoClass(typeof(DummyNamespaceWalk))]
internal interface INamespaceWalk
{
    void Walk(
        [MarshalAs(UnmanagedType.Interface)] object? punkToWalk,
        uint dwFlags,
        int cDepth,
        [MarshalAs(UnmanagedType.Interface)] INamespaceWalkCB pnswcb);
        
        void GetIDArrayResult(
            out uint pcItems,
            out IntPtr prgpidl); // The actual type is "PIDLIST_ABSOLUTE **"
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

    // Actual parameter types are "LPSTR*". If I write "out string" then runtime use it as "BStr", so I use "out IntPtr" here.
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

    // Actual parameter types are "LPSTR*". If I write "out string" then runtime use it as "BStr", so I use "out IntPtr" here.
    new void InitializeProgressDialog(
        out IntPtr ppszTitle,
        out IntPtr ppszCancel);

    void WalkComplete(int hr);
}
#pragma warning restore SYSLIB1096 // Convert to 'GeneratedComInterface'

// COMオブジェクト生成方法サンプル用途
[ComImport]
[Guid("3A9F4A2A-C7C6-4D9D-8F9E-D71EE470B57F")]
internal class DummyNamespaceWalk;
