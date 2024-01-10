# ComInteropTest

This is a project to test COM interop, especially when using `LibraryImportAttribute`

I wrote an article about this test, written in Japanese: [LibraryImportAttribute経由で取得したCOMオブジェクトのReleaseはほぼGC頼りになる](https://tan.hatenadiary.jp/entry/2024/01/11/004953)

## Summary
If we marshal a `ComObject` using `UniqueComInterfaceMarshaller<T>`, call `ComObject.FinalRelease`, and don't call `GC.SupressFinalize` for the `ComObject`, then **UNDEFINED BEHAVIOR** occurs. For example, `AccessViolationException` will occur.
