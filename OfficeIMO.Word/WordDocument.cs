using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Core;
using OfficeIMO.Shared;
using OfficeIMO.Word.Fluent;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides functionality for creating, loading and manipulating Word documents.
    /// </summary>
    public partial class WordDocument : IDisposable {
        internal int? _tableOfContentIndex;
        internal TableOfContentStyle? _tableOfContentStyle;
        private MemoryStream? _ownedPackageStream;
        private bool _tableOfContentUpdateQueued;
        private bool _disposed;
        private DocumentPersistenceMode _persistenceMode = DocumentPersistenceMode.Explicit;
    }
}
