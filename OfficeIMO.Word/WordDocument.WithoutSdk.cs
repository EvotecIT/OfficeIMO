using System.Diagnostics;
using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word {
    /// <summary>How a document loaded without the Open XML SDK opens the SDK package when it needs it.</summary>
    internal sealed class WordSdkDeferredLoad {
        internal WordSdkDeferredLoad(bool readOnly, OfficeOpenXmlLoadSettings? openSettings) {
            ReadOnly = readOnly;
            OpenSettings = openSettings;
        }

        internal bool ReadOnly { get; }
        internal OfficeOpenXmlLoadSettings? OpenSettings { get; }
    }

    /// <summary>
    /// A plain, unsigned Transitional document loads without opening the Open XML SDK. OfficeIMO's own Open XML layer
    /// holds the package, and paragraphs, runs, their text and basic formatting are read and edited on it. Anything else
    /// opens the SDK package through the bridge, built from the document as it stands, and from then on the SDK package
    /// is the document. Every such opening is recorded in <see cref="OpenXmlSdkFallbacks"/>.
    /// </summary>
    public partial class WordDocument {
        // Set while the SDK package has not been opened: the package and how the SDK opens when it is needed.
        internal WordPackageModel? _ownedModel;
        private WordSdkDeferredLoad? _sdkDeferredLoad;
        private readonly List<string> _openXmlSdkFallbacks = new();
        // Tests compare the two paths on the same input by turning the owned one off.
        internal static readonly AsyncLocal<bool> DisableOwnedOpenXmlForTesting = new();

        /// <summary>True while the document works on OfficeIMO's own Open XML layer and has not opened the SDK package.</summary>
        internal bool IsSdkDeferred => _ownedModel != null;

        /// <summary>
        /// Why the document opened the Open XML SDK package after a load that did not need it: one entry per opening,
        /// naming the member that needed it. Empty while the document works through OfficeIMO's own Open XML layer.
        /// </summary>
        public IReadOnlyList<string> OpenXmlSdkFallbacks => _openXmlSdkFallbacks;

        /// <summary>
        /// Loads <paramref name="sourceBytes"/> without the SDK when the package and the options allow it; returns null
        /// for the SDK load to handle exactly as before.
        /// </summary>
        private static WordDocument? TryLoadWithoutSdk(byte[] sourceBytes, WordLoadOptions options, string? filePath, Stream? originalStream) {
            if (DisableOwnedOpenXmlForTesting.Value || options.OverrideStyles && options.AccessMode != DocumentAccessMode.ReadOnly
                || options.FailOnLoss || !IsPlainOpenSettings(options.OpenSettings) || RequiresLegacyRuntimeSignatureSnapshot()) {
                return null;
            }
            WordPackageModel? model = WordPackageModel.TryOpen(sourceBytes, options.OpenSettings?.MaxCharactersInPart ?? 0);
            if (model == null) return null;
            bool readOnly = options.AccessMode == DocumentAccessMode.ReadOnly;
            return new WordDocument {
                FilePath = filePath,
                OriginalStream = originalStream,
                _persistenceMode = options.PersistenceMode,
                _ownedModel = model,
                _sdkDeferredLoad = new WordSdkDeferredLoad(readOnly, options.OpenSettings),
            };
        }

        private static bool IsPlainOpenSettings(OfficeOpenXmlLoadSettings? settings) =>
            settings == null
            || settings.CompatibilityLevel == new OfficeOpenXmlLoadSettings().CompatibilityLevel
            && settings.MarkupCompatibilityMode == new OfficeOpenXmlLoadSettings().MarkupCompatibilityMode
            && settings.MarkupCompatibilityTargetVersion == new OfficeOpenXmlLoadSettings().MarkupCompatibilityTargetVersion;

        /// <summary>
        /// Opens the SDK package over the document as it stands, through the bridge. Does nothing once it is open.
        /// </summary>
        internal void OpenSdkPackage(string? reason) {
            WordPackageModel? model = _ownedModel;
            if (model == null) return;
            WordSdkDeferredLoad load = _sdkDeferredLoad!;
            byte[] bytes = model.ToBytes();
            _ownedModel = null;
            _sdkDeferredLoad = null;
            _openXmlSdkFallbacks.Add(reason ?? DescribeSdkCaller());
            try {
                WordOpenXmlSdkBridge.OpenPackage(this, bytes, load);
            } finally {
                // Paragraphs handed out earlier find their SDK elements by their place in this tree, so it stays readable.
                if (model.IsDocumentParsed) model.Document.DetachFromPool();
                model.Dispose();
            }
        }

        /// <summary>Whether a document opened without the SDK is read-only.</summary>
        private bool IsOwnedReadOnly => _sdkDeferredLoad?.ReadOnly == true;

        /// <summary>
        /// Writes a document whose SDK package was never opened, when the save needs nothing the owned layer lacks.
        /// Returns false, having changed nothing, when the save has to go through the SDK.
        /// </summary>
        private bool TryCreateBytesWithoutSdk(string? filePath, WordSaveOptions? options, out byte[] bytes) {
            bytes = Array.Empty<byte>();
            WordPackageModel? model = _ownedModel;
            if (model == null || ResolveSavePassword(options) != null || options?.SignatureMutationPolicy != null || AutoUpdateToc) return false;
            // Other extensions change the document type, or the format, which the SDK save handles.
            if (!string.IsNullOrEmpty(filePath) && !string.Equals(Path.GetExtension(filePath!), ".docx", StringComparison.OrdinalIgnoreCase)) return false;

            if (WordOwnedSave.FindSaveBlocker(model) is string blocker) {
                OpenSdkPackage("Save: " + blocker);
                return false;
            }
            WordOwnedSave.PrepareForSave(model);
            if (!model.IsChanged) {
                bytes = model.SourceBytes;
                return true;
            }
            using var buffer = new MemoryStream(model.SourceBytes.Length + 4096);
            model.Save(buffer);
            bytes = buffer.ToArray();
            return true;
        }

        // Diagnostics only: names the first member outside the document's package accessors that needed the SDK. A
        // trimmed or NativeAOT image may describe its frames with less detail.
        private static string DescribeSdkCaller() {
            try {
                string trace = new StackTrace(2, false).ToString();
                foreach (string rawLine in trace.Split('\n')) {
                    string line = rawLine.Trim();
                    if (line.StartsWith("at ", StringComparison.Ordinal)) line = line.Substring(3);
                    int arguments = line.IndexOf('(');
                    if (arguments > 0) line = line.Substring(0, arguments);
                    if (line.Length == 0 || line.IndexOf("OfficeIMO.Word.", StringComparison.Ordinal) < 0) continue;
                    if (line.IndexOf("OpenSdkPackage", StringComparison.Ordinal) >= 0
                        || line.IndexOf("ResolveOwned", StringComparison.Ordinal) >= 0
                        || line.IndexOf(".get__", StringComparison.Ordinal) >= 0
                        || line.IndexOf("RequireSdk", StringComparison.Ordinal) >= 0) continue;
                    return line;
                }
            } catch (Exception exception) when (exception is not OutOfMemoryException) {
            }
            return "OfficeIMO.Word";
        }
    }
}
