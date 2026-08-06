using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using OfficeIMO.Core.Internal;
using OfficeIMO.Word.Fluent;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides functionality for creating, loading and manipulating Word documents.
    /// </summary>
    public partial class WordDocument : IDisposable, IAsyncDisposable {
        internal int? _tableOfContentIndex;
        internal WordTableOfContentsStyle? _tableOfContentStyle;
        private MemoryStream? _ownedPackageStream;
        // On .NET Framework, validation detaches the immutable encoded baseline from the
        // writable package stream so later unsaved edits can be flushed without corrupting it.
        private MemoryStream? _legacyValidationLivePackageStream;
        // System.IO.Packaging can rewrite the writable stream immediately after a part removal.
        // Retain the validated load-time bytes until validation detaches that live stream.
        private byte[]? _legacyValidationEncodedPackageBytes;
        private bool _tableOfContentUpdateQueued;
        private bool _disposed;
        private DocumentPersistenceMode _persistenceMode = DocumentPersistenceMode.Explicit;
    }
}
