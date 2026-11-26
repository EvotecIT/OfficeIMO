using System.Diagnostics;
using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word {
    /// <summary>How a document loaded without the Open XML SDK opens the SDK package when it needs it.</summary>
    internal sealed class WordSdkDeferredLoad {
        internal WordSdkDeferredLoad(bool readOnly, OfficeOpenXmlLoadSettings? openSettings, bool created = false) {
            ReadOnly = readOnly;
            OpenSettings = openSettings;
            Created = created;
        }

        internal bool ReadOnly { get; }
        internal OfficeOpenXmlLoadSettings? OpenSettings { get; }

        /// <summary>Whether the document was created from the default package, so the SDK opens it as a new document.</summary>
        internal bool Created { get; }
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
        // The package as it stood when a document that handed out paragraphs was disposed, and the copy parsed from it
        // on first use, so those paragraphs stay readable as they are on the SDK path.
        private byte[]? _disposedPackageBytes;
        private WordPackageModel? _disposedReadModel;
        // Where each element of the disposed tree sits in document order, and the copy's element at each position.
        private int[]? _disposedElementPositions;
        private int[]? _disposedElementsByPosition;
        private OfficeOpenXmlLoadSettings? _disposedOpenSettings;

        internal WordPackageModel? DisposedReadModel {
            get {
                if (_disposedReadModel == null && _disposedPackageBytes != null && _ownedNodeElements == null) {
                    _disposedReadModel = WordPackageModel.TryOpen(_disposedPackageBytes, 0);
                    if (_disposedReadModel == null) _disposedPackageBytes = null;
                    else _disposedElementsByPosition = _disposedReadModel.ElementsByPosition();
                }
                return _disposedReadModel;
            }
        }

        /// <summary>The copy's element that stands where <paramref name="node"/> stood in the disposed document, or NoNode.</summary>
        internal int DisposedCopyNode(int node) {
            if (node < 0 || _disposedElementPositions is not { } positions || _disposedElementsByPosition is not { } nodes) {
                return OfficeIMO.Core.Internal.OpenXml.OfficeXmlTree.NoNode;
            }
            int position = node < positions.Length ? positions[node] : -1;
            return position >= 0 && position < nodes.Length ? nodes[position] : OfficeIMO.Core.Internal.OpenXml.OfficeXmlTree.NoNode;
        }

        /// <summary>
        /// Opens an SDK package over the copy a disposed document keeps, so members its paragraphs need beyond the owned
        /// layer work after disposal as they do on the SDK path. The copy stays readable for paragraphs not yet resolved.
        /// </summary>
        internal void OpenSdkPackageAfterDispose() {
            WordPackageModel? copy = DisposedReadModel;
            if (copy == null || _ownedNodeElements != null) return;
            _ownedModel = copy;
            _sdkDeferredLoad = new WordSdkDeferredLoad(readOnly: true, _disposedOpenSettings);
            OpenSdkPackage("Paragraph used after the document was disposed", disposeModel: false);
        }

        // The owned tree the SDK package was built from, and the SDK element made from each of its nodes.
        internal OfficeIMO.Core.Internal.OpenXml.OfficeXmlTree? _ownedNodeTree;
        internal object?[]? _ownedNodeElements;
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

        /// <summary>
        /// Creates a document over <paramref name="defaultPackage"/>, the default package of a plain document with the default
        /// theme, without the SDK; returns null for the SDK path to create it exactly as before. The SDK opens over the same
        /// package later as it does for a new document.
        /// </summary>
        private static WordDocument? TryCreateWithoutSdk(string? filePath, Stream? stream, DocumentPersistenceMode persistenceMode,
            byte[] defaultPackage) {
            if (DisableOwnedOpenXmlForTesting.Value || RequiresLegacyRuntimeSignatureSnapshot()) return null;
            WordPackageModel? model = WordPackageModel.TryOpen(defaultPackage, 0);
            if (model == null) return null;
            return new WordDocument {
                FilePath = filePath,
                OriginalStream = stream,
                _persistenceMode = persistenceMode,
                _packageContentTypesAdvertiseXmlSignatures = false,
                _normalizeOpenOfficeRelationshipsOnSave = false,
                _tableNormalizationRequired = false,
                _creationTheme = OfficeTheme.Default,
                _ownedModel = model,
                _sdkDeferredLoad = new WordSdkDeferredLoad(readOnly: false, openSettings: null, created: true),
            };
        }

        /// <summary>
        /// The identifier to store for a built-in style, as <see cref="UseBuiltInStyle"/> returns it, when applying the
        /// style adds no definition; otherwise null, and the SDK path adds it.
        /// </summary>
        internal string? TryUseBuiltInStyleWithoutSdk(WordPackageModel model, string builtInStyleId) {
            WordOwnedStyles styles = model.Styles;
            string? styleId = styles.FindBuiltInStyleId(builtInStyleId);
            if (styleId != null && !string.Equals(styleId, builtInStyleId, StringComparison.Ordinal)) return styleId;
            if (FileOpenAccess == FileAccess.Read) return builtInStyleId;
            return styles.DefinesWithDependencies(builtInStyleId) ? builtInStyleId : null;
        }

        /// <summary>
        /// Does what <see cref="HeadingModified"/> does after a heading style is applied, when that is nothing: a settings
        /// part without a field update requested on open, and no content control or field in the body that could be a table
        /// of contents. Returns false for the SDK path to handle it.
        /// </summary>
        internal bool TryHeadingModifiedWithoutSdk(WordPackageModel model) {
            if (_deferHeadingModified) {
                _headingModifiedWhileDeferred = true;
                return true;
            }
            // Reading the settings on the SDK path adds a settings part to a package without one.
            if (_tableOfContentUpdateQueued || !model.HasSettings || model.SettingsHave("updateFields")) return false;
            // The SDK path finds a table of contents in a content control or in a TOC field of a body paragraph.
            return !model.BodyHasFieldsOrContentControls;
        }

        private static bool IsPlainOpenSettings(OfficeOpenXmlLoadSettings? settings) =>
            settings == null
            || settings.CompatibilityLevel == new OfficeOpenXmlLoadSettings().CompatibilityLevel
            && settings.MarkupCompatibilityMode == new OfficeOpenXmlLoadSettings().MarkupCompatibilityMode
            && settings.MarkupCompatibilityTargetVersion == new OfficeOpenXmlLoadSettings().MarkupCompatibilityTargetVersion;

        /// <summary>
        /// Opens the SDK package over the document as it stands, through the bridge. Does nothing once it is open.
        /// </summary>
        internal void OpenSdkPackage(string? reason) => OpenSdkPackage(reason, disposeModel: true);

        private void OpenSdkPackage(string? reason, bool disposeModel) {
            WordPackageModel? model = _ownedModel;
            if (model == null) return;
            WordSdkDeferredLoad load = _sdkDeferredLoad!;
            byte[] bytes = model.ToBytes();
            _ownedModel = null;
            _sdkDeferredLoad = null;
            try {
                WordOpenXmlSdkBridge.OpenPackage(this, bytes, load);
            } catch {
                // The document keeps working on the owned layer, and the next member that needs the SDK tries again.
                _ownedModel = model;
                _sdkDeferredLoad = load;
                throw;
            }
            _openXmlSdkFallbacks.Add(reason ?? DescribeSdkCaller());
            // Paragraphs handed out earlier resolve through this map, which is built before anything can edit the SDK tree.
            if (model.IsDocumentParsed) {
                _ownedNodeTree = model.Document;
                _ownedNodeElements = WordOpenXmlSdkBridge.MapOwnedTree(model.Document, _documentValue!);
            }
            if (disposeModel) model.Dispose();
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
            model.LastSavedBytes = bytes;
            model.LastSavedVersion = model.Document.Version;
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
