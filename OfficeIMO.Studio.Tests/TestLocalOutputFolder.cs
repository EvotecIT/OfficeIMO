using System.Reflection;
using Avalonia.Platform.Storage;

namespace OfficeIMO.Studio.Tests;

/// <summary>A selected local folder boundary; selecting it never creates child items.</summary>
internal static class TestLocalOutputFolder {
    internal static IStorageFolder Create(string path) {
        var item = DispatchProxy.Create<IStorageFolder, TestStorageFile.StorageProxy>();
        ((TestStorageFile.StorageProxy)(object)item).Call = (method, _) => method switch {
            "get_Name" => Path.GetFileName(path),
            "get_Path" => new Uri(Path.GetFullPath(path)),
            "get_CanBookmark" => false,
            "Dispose" => null,
            _ => throw new NotSupportedException(method)
        };
        return item;
    }
}
