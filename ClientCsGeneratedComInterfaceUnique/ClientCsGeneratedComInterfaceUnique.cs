using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using System.Runtime.Versioning;

namespace ClientTestCSharpComImport;

[SupportedOSPlatform("windows")]

internal class ClientCsGeneratedComInterfaceUnique
{
    [STAThread]
    private static void Main()
    {
        UseCom();
        GC.Collect();

        // ComObject.FinalReleaseを呼び出していない場合は、ComObject.Finalizeが呼びだされて各種ネイティブCOMオブジェクトはReleaseされます。
        // ComObject.FinalReleaseを呼び出しており、かつComObject.Finalizeも実行する場合は、F11でステップ実行すると「System.AccessViolationException: 'Attempted to read or write protected memory. This is often an indication that other memory is corrupt.'」例外が発生する場合や「System.NullReferenceException: 'Object reference not set to an instance of an object.'」が発生する場合がありました。
        // ただし、F10等のステップ実行やF5でまとめて実行すると、このコードでは終了コード0で完了しました。
        // もちろん、ComObject.FinalReleaseを呼び出しており、かつGC.SupressFinalizeで抑制している場合は、既にRelease済みかつGC中でも何も起こらないので問題ありません。
        GC.WaitForPendingFinalizers();

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
            out object objNamespaceWalk));
        var pNamespaceWalk = (INamespaceWalk)objNamespaceWalk;
        Marshal.ThrowExceptionForHR(NativeMethods.CoCreateInstance(
            NativeMethods.CLSID_DummyNamespaceWalkCB,
            null,
            NativeMethods.CLSCTX_INPROC_SERVER,
            typeof(INamespaceWalkCB).GUID,
            out object objNamespaceWalkCB));
        var pNamespaceWalkCB = (INamespaceWalkCB)objNamespaceWalkCB;
        var pNamespaceWalkCB2 = (INamespaceWalkCB2)pNamespaceWalkCB;

        // 適当に使う
        pNamespaceWalkCB.InitializeProgressDialog(out _, out _);
        pNamespaceWalkCB2.WalkComplete(0);
        pNamespaceWalk.Walk(null, 0, 0, pNamespaceWalkCB);

        // 使い終わったので解放
        Console.Write("""
            Do you test to what hapen if you explicitly call ComObject.FinalRelease?
            0: Don't call ComObject.FinalRelease.
            1: Call ComObject.FinalRelease only. ComObject.Finalize will be called later, and cause UNDEFINED BEHAVIOR!.
            2: Call ComObject.FinalRelease and GC.SuppressFinalize.
            > 
            """);
        _ = int.TryParse(Console.ReadLine(), out int choise);
        if (choise >= 1)
        {
            // GeneratedComInterface用のinterfaceはMarshal.ReleaseComObjectへ渡せません。渡すとArgumentExceptionになります。
            // UniqueComInterfaceMarshallerを使う場合だと、ComObject.UniqueInstanceがtrueであるため、ComObject.FinalRelease()で明示的に解放できます。
            ((ComObject)objNamespaceWalkCB).FinalRelease();
            ((ComObject)objNamespaceWalk).FinalRelease();

            // Release後に使おうとすると、Use-After-Freeが起こります。
            // 試すと、「System.AccessViolationException: 'Attempted to read or write protected memory. This is often an indication that other memory is corrupt.'」例外が発生しました。
            // pNamespaceWalk.Walk(null, 0, 0, pNamespaceWalkCB);

            if (choise >= 2)
            {
                // このままではGC発動時のFinalize時に改めてReleaseしようとしてUse-After-Freeが起こります。そういうわけで防ぎます。
#pragma warning disable CA1816 // Dispose methods should call SuppressFinalize
                GC.SuppressFinalize(objNamespaceWalk);
                GC.SuppressFinalize(objNamespaceWalkCB);
#pragma warning restore CA1816 // Dispose methods should call SuppressFinalize
            }
        }
    }
}

internal static partial class NativeMethods
{
    // 最後のout引数にUniqueComInterfaceMarshallerを指定しています
    [LibraryImport("Ole32.dll")]
    internal static partial int CoCreateInstance(
        in Guid rclsid,
        [MarshalAs(UnmanagedType.Interface)] object? pUnkOuter,
        uint dwClsContext,
        in Guid riid,
        [MarshalUsing(typeof(UniqueComInterfaceMarshaller<object>))] out object ppv);

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
