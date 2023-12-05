using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using System.Runtime.Versioning;

namespace ClientTestCSharpComImport;

[SupportedOSPlatform("windows")]

internal class TestGeneratedComInterfaceUnique
{
    private static void Main()
    {
        UseCom();
        GC.Collect();
        GC.WaitForPendingFinalizers(); // 今回はComObject.FinalRelease()で解放できているので、ファイナライザを待つ必要はありません
        GC.Collect();
    }

    private static void UseCom()
    {
        Marshal.ThrowExceptionForHR(NativeMethods.CoCreateInstance(
            NativeMethods.CLSID_DummyNamespaceWalk,
            null,
            NativeMethods.CLSCTX_LOCAL_SERVER,
            typeof(INamespaceWalk).GUID,
            out object objNamespaceWalk));
        var pNamespaceWalk = (INamespaceWalk)objNamespaceWalk;
        Marshal.ThrowExceptionForHR(NativeMethods.CoCreateInstance(
            NativeMethods.CLSID_DummyNamespaceWalkCB,
            null,
            NativeMethods.CLSCTX_LOCAL_SERVER,
            typeof(INamespaceWalkCB).GUID,
            out object objNamespaceWalkCB));
        var pNamespaceWalkCB = (INamespaceWalkCB)objNamespaceWalkCB;
        var pNamespaceWalkCB2 = (INamespaceWalkCB2)pNamespaceWalkCB;

        pNamespaceWalkCB.InitializeProgressDialog(out _, out _);
        pNamespaceWalkCB2.WalkComplete(0);
        pNamespaceWalk.Walk(null, 0, 0, pNamespaceWalkCB);

        // GeneratedComInterface用のinterfaceはMarshal.ReleaseComObjectへ渡せません。渡すとArgumentExceptionになります。
        // UniqueComInterfaceMarshallerを使う場合だと、ComObject.UniqueInstance==trueであるため、ComObject.FinalRelease()を呼び出した時点で解放できます！
        ((ComObject)objNamespaceWalkCB).FinalRelease();
        ((ComObject)objNamespaceWalk).FinalRelease();

        // ただそのままだとファイナライザでも改めてReleaseしようとします。Use-After-Freeでは？そういうわけで防ぎます。
        GC.SuppressFinalize(objNamespaceWalk);
        GC.SuppressFinalize(objNamespaceWalkCB);

        // 解放後に使おうとすると、Use-After-Freeかも……？
        // pNamespaceWalk.Walk(null, 0, 0, pNamespaceWalkCB);
        ((ComObject)objNamespaceWalk).FinalRelease();
        ((ComObject)objNamespaceWalk).FinalRelease();
        ((ComObject)objNamespaceWalk).FinalRelease();
        ((ComObject)objNamespaceWalk).FinalRelease();
        ((ComObject)objNamespaceWalk).FinalRelease();
        ((ComObject)objNamespaceWalk).FinalRelease();
        ((ComObject)objNamespaceWalk).FinalRelease();
        ((ComObject)objNamespaceWalk).FinalRelease();
        ((ComObject)objNamespaceWalk).FinalRelease();
        ((ComObject)objNamespaceWalk).FinalRelease();
        ((ComObject)objNamespaceWalk).FinalRelease();
        ((ComObject)objNamespaceWalk).FinalRelease();
        ((ComObject)objNamespaceWalk).FinalRelease();
        ((ComObject)objNamespaceWalk).FinalRelease();
        ((ComObject)objNamespaceWalk).FinalRelease();
        ((ComObject)objNamespaceWalk).FinalRelease();
    }
}

internal static partial class NativeMethods
{
    // 最後の引数にUniqueComInterfaceMarshallerを追加しています！！！！！！！！
    [LibraryImport("Ole32.dll")]
    internal static partial int CoCreateInstance(
        in Guid rclsid,
        [MarshalAs(UnmanagedType.Interface)] object? pUnkOuter,
        uint dwClsContext,
        in Guid riid,
        [MarshalUsing(typeof(UniqueComInterfaceMarshaller<object>))] out object ppv);

    internal const uint CLSCTX_LOCAL_SERVER = 4;

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
        out IntPtr prgpidl); // actual is "PIDLIST_ABSOLUTE **"
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

    // 引数は実際はLPWSTR *だけれど、outで受けるとBStrとして扱いそうなのでIntPtrでごまかす
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
