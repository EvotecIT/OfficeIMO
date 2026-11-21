using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Word;

public partial class WordDocument {
    /// <summary>Reopens snapshot bytes without resetting the source package's XML limits or compatibility policy.</summary>
    internal static WordDocument LoadPackageSnapshot(Stream snapshot, WordprocessingDocument source) =>
        LoadFromStream(snapshot, new WordLoadOptions {
            MaxInputBytes = Math.Max(WordLoadOptions.DefaultMaxInputBytes, snapshot.Length)
        }, captureSourcePackageSecurity: false, detectXml: false,
            packageSettings: WordPackageSnapshot.CreateOpenSettings(source, autoSave: false));
}
