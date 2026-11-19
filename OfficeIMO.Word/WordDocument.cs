using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using OfficeIMO.Core.Internal;
using System.Text.RegularExpressions;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides functionality for creating, loading and manipulating Word documents.
    /// </summary>
    /// <example>
    /// <code language="csharp">
    /// using OfficeIMO.Word;
    ///
    /// using var document = WordDocument.Create();
    /// document.AddParagraph("Hello from OfficeIMO.");
    /// document.Save("hello.docx");
    /// </code>
    /// </example>
    public partial class WordDocument : IDisposable, IAsyncDisposable {
        internal int? _tableOfContentIndex;
        internal WordTableOfContentsStyle? _tableOfContentStyle;
        // The table of contents this instance created, refreshed on save; loaded TOCs keep their stored entries.
        internal WordTableOfContent? _generatedTableOfContent;
        private MemoryStream? _ownedPackageStream;
        // On .NET Framework, validation detaches the immutable encoded baseline from the
        // writable package stream so later unsaved edits can be flushed without corrupting it.
        private MemoryStream? _legacyValidationLivePackageStream;
        // System.IO.Packaging can rewrite the writable stream immediately after a part removal.
        // Retain the validated load-time bytes until validation detaches that live stream.
        private byte[]? _legacyValidationEncodedPackageBytes;
        private bool _tableOfContentUpdateQueued;
        private bool _tableOfContentEnabledUpdateFieldsOnOpen;
        private bool _tableNormalizationRequired = true;
        private bool _normalizeOpenOfficeRelationshipsOnSave = true;
        private bool _disposed;
        internal bool _openXmlDocumentAccessed;
        /// <summary>The style catalog last read for this document; see <see cref="WordStyleCatalog.For"/>.</summary>
        internal WordStyleCatalog? _styleCatalog;
        /// <summary>Changes whenever OfficeIMO replaces the styles part's content without loading its DOM.</summary>
        internal int _styleDefinitionsRevision;
        private DocumentPersistenceMode _persistenceMode = DocumentPersistenceMode.Explicit;
    }
}
