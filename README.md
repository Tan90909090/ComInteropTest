# ComInteropTest

This is a project to test COM interop, especially when using `LibraryImportAttribute`

I wrote an article about this test, written in Japanese: [LibraryImportAttribute経由で取得したCOMオブジェクトのReleaseはほぼGC頼りになる](https://tan.hatenadiary.jp/entry/2024/01/11/004953)

## Summary
In .NET 8, if we marshal a `ComObject` using `UniqueComInterfaceMarshaller<T>`, call `ComObject.FinalRelease`, and don't call `GC.SupressFinalize` for the `ComObject`, then **UNDEFINED BEHAVIOR** occurs. For example, `AccessViolationException` will occur.

In .NET 9, **this issue will be fixed**: [Make multiple calls to FinalRelease safe by jkoritzinsky · Pull Request #97059 · dotnet/runtime](https://github.com/dotnet/runtime/pull/97059)