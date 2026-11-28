using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Word.OpenXml;
using System.Text.RegularExpressions;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        /// <summary>
        /// Collection of all comments inserted into the document.
        /// </summary>
        /// <returns>List of <see cref="WordComment"/> objects for each comment.</returns>
        public List<WordComment> Comments {
            get { return WordComment.GetAllComments(this); }
        }

        /// <summary>
        /// Removes comment with the specified id.
        /// </summary>
        /// <param name="commentId">Id of the comment to remove.</param>
        public void RemoveComment(string commentId) {
            var comment = this.Comments.FirstOrDefault(c => c.Id == commentId);
            comment?.Remove();
        }

        /// <summary>
        /// Removes the specified comment from the document.
        /// </summary>
        /// <param name="comment">Comment instance to remove.</param>
        public void RemoveComment(WordComment comment) {
            comment?.Remove();
        }

        /// <summary>
        /// Removes all comments from the document.
        /// </summary>
        public void RemoveAllComments() {
            foreach (var comment in this.Comments.ToList()) {
                comment.Remove();
            }
        }

        /// <summary>
        /// Gets the value of a document variable or <c>null</c> if the variable does not exist.
        /// </summary>
        /// <param name="name">Variable name.</param>
        public string? GetDocumentVariable(string name) {
            return DocumentVariables.TryGetValue(name, out var value) ? value : null;
        }

        /// <summary>
        /// Sets the value of a document variable. Creates it if it does not exist.
        /// </summary>
        /// <param name="name">Variable name.</param>
        /// <param name="value">Variable value.</param>
        public void SetDocumentVariable(string name, string value) {
            DocumentVariables[name] = value;
        }

        /// <summary>
        /// Removes the document variable with the specified name if present.
        /// </summary>
        /// <param name="name">Variable name.</param>
        public void RemoveDocumentVariable(string name) {
            DocumentVariables.Remove(name);
        }

        /// <summary>
        /// Removes the document variable at the specified index.
        /// </summary>
        /// <param name="index">Zero-based index of the variable to remove.</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when index is out of range.</exception>
        public void RemoveDocumentVariableAt(int index) {
            ArgumentOutOfRangeException.ThrowIfNegative(index);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(index, DocumentVariables.Count);
            string key = DocumentVariables.Keys.ElementAt(index);
            DocumentVariables.Remove(key);
        }

        /// <summary>
        /// Determines whether the document contains any document variables.
        /// </summary>
        public bool HasDocumentVariables => DocumentVariables.Count > 0;

        /// <summary>
        /// Returns a read-only view of all document variables.
        /// </summary>
        public IReadOnlyDictionary<string, string> GetDocumentVariables() {
            return new Dictionary<string, string>(DocumentVariables);
        }

        /// <summary>
        /// Enable or disable tracking of all revisions, moves and formatting changes.
        /// </summary>
        public bool TrackChanges {
            get => this.Settings.TrackRevisions;
            set {
                this.Settings.TrackRevisions = value;
                this.Settings.TrackFormatting = value;
                this.Settings.TrackMoves = value;
            }
        }

        /// <summary>
        /// Collection of sections contained in the document.
        /// </summary>
        public List<WordSection> Sections {
            get {
                if (_ownedModel is { } model) return _ownedSections ??= WordSection.LoadOwnedSections(this, model);
                if (IsDisposed && _wordprocessingDocumentValue == null && _ownedSections != null) return _ownedSections;
                return _sections;
            }
        }

        private List<WordSection> _sections = new List<WordSection>();

        /// <summary>
        /// Path to the file backing this document.
        /// </summary>
        public string? FilePath { get; internal set; }

        /// <summary>
        /// Original stream where this document was created / loaded from.
        /// </summary>
        internal Stream? OriginalStream { get; set; }

        private WordSettings? _settings;

        /// <summary>
        /// Provides access to document settings. A missing settings part is created on first use.
        /// </summary>
        public WordSettings Settings {
            get => _settings ??= new WordSettings(this);
            internal set => _settings = value;
        }

        internal OfficeXmlNode GetSettingsRoot(bool create = false) {
            if (_ownedModel is { } model) {
                OfficeXmlNode root = model.SettingsRoot(create: false);
                if (root.Exists || !create) return root;
                if (IsOwnedReadOnly) throw new IOException("Cannot modify readonly container");
                return model.SettingsRoot(create: true);
            }
            MainDocumentPart? main = _wordprocessingDocumentValue?.MainDocumentPart;
            DocumentSettingsPart? part = main?.DocumentSettingsPart;
            if (create && main != null) {
                part ??= main.AddNewPart<DocumentSettingsPart>();
                part.Settings ??= new Settings();
            }
            return OfficeXmlNode.FromSdk(part?.Settings);
        }

        internal OfficeXmlNode GetDefaultRunProperties(bool edit) {
            OfficeXmlNode defaults;
            if (_ownedModel is { } model) {
                OfficeXmlNode styles = model.StylesRoot(create: edit && !IsOwnedReadOnly);
                defaults = edit && styles.Exists
                    ? styles.GetOrPrependChild(WordXml.W, "docDefaults")
                    : styles.WChild("docDefaults");
            } else if (edit) {
                EnsureStyleDefinitionsInitialized();
                OfficeXmlNode styles = OfficeXmlNode.FromSdk(_wordprocessingDocumentValue?.MainDocumentPart?.StyleDefinitionsPart?.Styles);
                defaults = styles.IsNone ? default : styles.GetOrPrependChild(WordXml.W, "docDefaults");
            } else {
                defaults = WordStyleCatalog.ReadDocumentDefaults(_wordprocessingDocumentValue?.MainDocumentPart?.StyleDefinitionsPart);
            }
            if (!edit || defaults.IsNone) return defaults.WChild("rPrDefault").WChild("rPr");
            return defaults.GetOrPrependChild(WordXml.W, "rPrDefault").GetOrPrependChild(WordXml.W, "rPr");
        }

        internal OfficeXmlNode SettingsDocumentRoot => _ownedModel is { } model
            ? new OfficeXmlNode(model.Document, model.Document.Root)
            : OfficeXmlNode.FromSdk(_wordprocessingDocumentValue?.MainDocumentPart?.Document);

        /// <summary>
        /// Manages application related properties.
        /// </summary>
        public WordApplicationProperties ApplicationProperties {
            get {
                if (_ownedModel != null) OpenSdkPackage(null);
                return _applicationProperties;
            }
            internal set => _applicationProperties = value;
        }

        private WordApplicationProperties _applicationProperties = null!;

        /// <summary>
        /// Provides access to built-in document properties.
        /// </summary>
        public WordBuiltinDocumentProperties Properties {
            get {
                if (_ownedModel != null) OpenSdkPackage(null);
                return _properties;
            }
            internal set => _properties = value;
        }

        private WordBuiltinDocumentProperties _properties = null!;

        private WordCoverPageProperties? _coverPageProperties;

        /// <summary>
        /// Provides access to the cover page properties custom XML part used by built-in templates.
        /// </summary>
        public WordCoverPageProperties CoverPageProperties => _coverPageProperties ??= new WordCoverPageProperties(this);

        /// <summary>
        /// Collection of custom document properties.
        /// </summary>
        public Dictionary<string, WordCustomProperty> CustomProperties {
            get {
                if (_ownedModel != null) OpenSdkPackage(null);
                return _customProperties;
            }
        }

        private readonly Dictionary<string, WordCustomProperty> _customProperties = new Dictionary<string, WordCustomProperty>();
        /// <summary>
        /// Collection of document variables accessible via <see cref="WordFieldType.DocVariable"/> fields.
        /// </summary>
        public Dictionary<string, string> DocumentVariables {
            get {
                if (_ownedModel != null) OpenSdkPackage(null);
                return _documentVariables;
            }
        }

        private readonly Dictionary<string, string> _documentVariables = new Dictionary<string, string>();

        /// <summary>
        /// Collection of bibliographic sources used in the document.
        /// </summary>
        public Dictionary<string, WordBibliographySource> BibliographySources {
            get {
                if (_ownedModel != null) OpenSdkPackage(null);
                return _bibliographySources;
            }
        }

        private readonly Dictionary<string, WordBibliographySource> _bibliographySources = new Dictionary<string, WordBibliographySource>();

        /// <summary>
        /// Provides basic statistics for the document.
        /// </summary>
        public WordDocumentStatistics Statistics {
            get {
                if (_ownedModel != null) OpenSdkPackage(null);
                return _statistics;
            }
            internal set => _statistics = value;
        }

        private WordDocumentStatistics _statistics = null!;

        /// <summary>
        /// Gets the persistence policy selected when the document was created or loaded.
        /// </summary>
        public DocumentPersistenceMode PersistenceMode => _persistenceMode;

        /// <summary>
        /// When <c>true</c> the table of contents is flagged to update before saving.
        /// </summary>
        public bool AutoUpdateToc { get; set; }


        private WordprocessingDocument? _wordprocessingDocumentValue;

        /// <summary>The SDK package. A document loaded without it opens it here, through the bridge, on first use.</summary>
        internal WordprocessingDocument _wordprocessingDocument {
            get {
                if (_ownedModel != null) OpenSdkPackage(null);
                return _wordprocessingDocumentValue!;
            }
            set => _wordprocessingDocumentValue = value;
        }

        /// <summary>
        /// Root document element.
        /// </summary>
        internal Document _document {
            get {
                if (_ownedModel != null) OpenSdkPackage(null);
                return _documentValue!;
            }
            set => _documentValue = value;
        }

        private Document? _documentValue;

        /// <summary>Gets the underlying Open XML package for advanced integration scenarios.</summary>
        public WordprocessingDocument OpenXmlDocument {
            get {
                // Once the mutable package escapes, Open XML edits cannot be observed reliably.
                // Resolve against current parts instead of reusing a derived numbering catalog.
                OpenSdkPackage("WordDocument.OpenXmlDocument");
                _openXmlDocumentAccessed = true;
                // Identifiers are rescanned on every allocation from now on.
                WordDrawingIdAllocator.MarkPackageShared(_wordprocessingDocument);
                WordStyleDefinitionsReader.StopStreaming(_wordprocessingDocument);
                return _wordprocessingDocument;
            }
        }


        /// <summary>
        /// FileOpenAccess of the document
        /// </summary>
        public FileAccess FileOpenAccess => _ownedModel != null
            ? IsOwnedReadOnly ? FileAccess.Read : FileAccess.ReadWrite
            : _wordprocessingDocument.FileOpenAccess;

        /// <summary>Gets whether the document was loaded for reading or editing.</summary>
        public DocumentAccessMode AccessMode => FileOpenAccess == FileAccess.Read
            ? DocumentAccessMode.ReadOnly
            : DocumentAccessMode.ReadWrite;
    }
}
