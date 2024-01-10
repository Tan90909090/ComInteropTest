using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using System.Runtime.Versioning;

namespace ClientTestCSharpComImport;

[SupportedOSPlatform("windows")]

internal class ClientCsGeneratedComInterface
{
    [STAThread]
    private static void Main()
    {
        UseCom();
        GC.Collect();
        GC.WaitForPendingFinalizers(); // ComObject.Finalizeが呼び出されて各種ネイティブCOMオブジェクトはReleaseされます
        GC.Collect();
    }

    private static void UseCom()
    {
        // COMオブジェクト生成
        Marshal.ThrowExceptionForHR(NativeMethods.CoCreateInstance(
            NativeMethods.CLSID_DummyNamespaceWalk,
            null,
            NativeMethods.CLSCTX_INPROC_SERVER,
            typeof(INamespaceWalk).GUID,
            out object? objNamespaceWalk));
        var pNamespaceWalk = (INamespaceWalk)objNamespaceWalk;

        Marshal.ThrowExceptionForHR(NativeMethods.CoCreateInstance(
            NativeMethods.CLSID_DummyNamespaceWalkCB,
            null,
            NativeMethods.CLSCTX_INPROC_SERVER,
            typeof(INamespaceWalkCB).GUID,
            out object? objNamespaceWalkCB));
        var pNamespaceWalkCB = (INamespaceWalkCB)objNamespaceWalkCB;
        var pNamespaceWalkCB2 = (INamespaceWalkCB2)pNamespaceWalkCB;

        // 適当に使う
        pNamespaceWalkCB.InitializeProgressDialog(out _, out _);
        pNamespaceWalkCB2.WalkComplete(0);
        pNamespaceWalk.Walk(null, 0, 0, pNamespaceWalkCB);

        // 使い終わったので解放したかったです
        // LibraryImportAttributeで生成したCOMオブジェクトはMarshal.ReleaseComObjectへ渡せません。渡すとArgumentExceptionになります。
        // Marshal.ReleaseComObject(pNamespaceWalk); // 「SYSLIB1099 The method 'int Marshal.ReleaseComObject(object o)' only supports runtime-based COM interop and will not work with type 'INamespaceWalk'」

        // また、デフォルトのComInterfaceMarshaller<T>の場合だとComObject.UniqueInstanceがfalseであるため、ComObject.FinalRelease()を呼んでも無意味です
        ((ComObject)objNamespaceWalkCB).FinalRelease(); // 無意味、Releaseされない
        ((ComObject)objNamespaceWalk).FinalRelease(); // 無意味、Releaseされない

        // Releaseされていないので、COMオブジェクトをまだ使えます
        pNamespaceWalk.Walk(null, 0, 0, pNamespaceWalkCB);
    }
}

internal static partial class NativeMethods
{
    [LibraryImport("Ole32.dll")]
    internal static partial int CoCreateInstance(
        in Guid rclsid,
        [MarshalAs(UnmanagedType.Interface)] object? pUnkOuter,
        uint dwClsContext,
        in Guid riid,
        [MarshalAs(UnmanagedType.Interface)] out object ppv);

    internal const uint CLSCTX_INPROC_SERVER = 1;

    internal static readonly Guid CLSID_DummyNamespaceWalk = Guid.Parse("{3A9F4A2A-C7C6-4D9D-8F9E-D71EE470B57F}");
    internal static readonly Guid CLSID_DummyNamespaceWalkCB = Guid.Parse("{1B94A81D-A105-428D-8902-849F8D6483A2}");
}

[GeneratedComInterface]
[Guid("57ced8a7-3f4a-432c-9350-30f24483f74f")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal partial interface INamespaceWalk
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

[GeneratedComInterface]
[Guid("d92995f8-cf5e-4a76-bf59-ead39ea2b97e")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal partial interface INamespaceWalkCB
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

// GeneratedComInterface使用時は、基底インターフェース内容のコピペは不要
[GeneratedComInterface]
[Guid("7ac7492b-c38e-438a-87db-68737844ff70")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal partial interface INamespaceWalkCB2 : INamespaceWalkCB
{
    void WalkComplete(int hr);
}
