using System.Runtime.ExceptionServices;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Drawing;
using OfficeIMO.Core.Internal;
using A = DocumentFormat.OpenXml.Drawing;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a PowerPoint presentation providing create, load, and save operations.
    /// </summary>
    public sealed partial class PowerPointDocument : IDisposable, IAsyncDisposable {
        // The Open XML SDK package and the slides built from it. A presentation loaded from an Open XML package opens
        // them on first use, so loading and saving an unchanged deck never needs them.
        private PresentationDocument? _sdkDocument;
        private PresentationPart? _sdkPresentationPart;
        private readonly List<PowerPointSlide> _sdkSlides = new();
        private DocumentAccessMode _accessMode;
        private PowerPointBuiltinDocumentProperties? _properties;
        private PowerPointApplicationProperties? _applicationProperties;
        private long _slideRelationshipSearchStart = 1;
        private string _filePath;
        private Stream? _packageStream;
        private Stream? _sourceStream;
        private DocumentPersistenceMode _persistenceMode = DocumentPersistenceMode.Explicit;
        private long _maximumSlideXmlCharacters = PowerPointXmlReader.MaximumPackageXmlCharacters;
        private bool _discardChangesOnDispose;
        private string? _signedPackageOpenFingerprint;
        private PowerPointSlideSize? _slideSize;
        private bool _disposed = false;
        private const int StreamBufferSize = 4096;
        private const string P14Namespace = "http://schemas.microsoft.com/office/powerpoint/2010/main";
        private const string SectionListUri = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}";
        private const string DefaultSectionName = "Section 1";
        private const string TableStylesResourceName = "OfficeIMO.PowerPoint.Resources.tableStyles.xml";
        private PowerPointDocument(PresentationDocument document, string filePath, bool isNewDocument) {
            _filePath = filePath;
            InitializeFromSdk(document, isNewDocument);
        }

        /// <summary>Whether the Open XML SDK package has been opened for this presentation.</summary>
        internal bool IsSdkLoaded => _sdkDocument != null;

        private PresentationDocument SdkDocument => EnsureSdk();

        // The package's document type: from the SDK once it is open, otherwise from the main part's content type.
        private PresentationDocumentType? _loadedDocumentType;

        private PresentationDocumentType? OpenXmlDocumentType => _sdkDocument?.DocumentType ?? _loadedDocumentType;

        private PresentationPart SdkPresentationPart {
            get {
                EnsureSdk();
                return _sdkPresentationPart!;
            }
        }

        private List<PowerPointSlide> LoadedSlides {
            get {
                EnsureSdk();
                return _sdkSlides;
            }
        }

        private PresentationDocument EnsureSdk() {
            if (_sdkDocument != null) return _sdkDocument;
            if (_disposed) throw new ObjectDisposedException(nameof(PowerPointDocument));
            if (_sdkDeferred) return OpenDeferredSdk();
            throw new InvalidOperationException("The presentation has no Open XML package to open.");
        }

        private void InitializeFromSdk(PresentationDocument document, bool isNewDocument) {
            _sdkDocument = document;
            _accessMode = document.FileOpenAccess == FileAccess.Read ? DocumentAccessMode.ReadOnly : DocumentAccessMode.ReadWrite;
            _sdkPresentationPart = document.PresentationPart ?? document.AddPresentationPart();
            if (isNewDocument || SdkPresentationPart.Presentation == null) {
                // New presentation - create with required initial structure
                PresentationRoot = new Presentation();
                InitializeDefaultParts();

                // InitializeDefaultParts creates the native master/layout scaffolding and one temporary slide.
                if (PresentationRoot.SlideIdList != null) {
                    foreach (SlideId slideId in PresentationRoot.SlideIdList.Elements<SlideId>()) {
                        string? relId = PowerPointUtils.GetRelationshipIdValue(slideId);
                        if (!string.IsNullOrEmpty(relId)) {
                            SlidePart slidePart = (SlidePart)SdkPresentationPart.GetPartById(relId!);
                            LoadedSlides.Add(new PowerPointSlide(slidePart, this));
                        }
                    }
                }
                if (isNewDocument) {
                    ClearSlides();
                }
            } else {
                // Loading existing presentation
                LoadExistingSlides();
            }

            document.Features.Set<PowerPointDocument>(this);
        }

        /// <summary>Whether the shared chart axis seed already covers this package's charts.</summary>
        internal bool ChartAxisSeedApplied { get; set; }

        private static byte[] ReadAllBytes(Stream stream) {
            return OfficeStreamReader.ReadAllBytes(stream);
        }

        /// <summary>Gets the destination path associated with the presentation, if any.</summary>
        public string? FilePath => string.IsNullOrEmpty(_filePath) ? null : _filePath;

        /// <summary>Gets the underlying Open XML package for advanced integration scenarios.</summary>
        public PresentationDocument OpenXmlDocument {
            get {
                ThrowIfDisposed();
                return SdkDocument!;
            }
        }

        /// <summary>Gets the configured persistence behavior.</summary>
        public DocumentPersistenceMode PersistenceMode => _persistenceMode;

        /// <summary>Gets whether the presentation is editable or read-only.</summary>
        public DocumentAccessMode AccessMode {
            get {
                ThrowIfDisposed();
                return _accessMode;
            }
        }

        /// <summary>
        ///     Collection of slides in the presentation.
        /// </summary>
        public IReadOnlyList<PowerPointSlide> Slides {
            get {
                ThrowIfDisposed();
                // Until a slide is needed, the count comes from the slide list the load read.
                return _sdkDeferred && _sdkDocument == null ? _slideView ??= new DeferredSlideList(this) : LoadedSlides;
            }
        }

        private DeferredSlideList? _slideView;

        /// <summary>The slides of a presentation whose SDK package is not open yet, read from its package; when the
        /// package can't be read without the SDK, reading a slide opens it.</summary>
        private sealed class DeferredSlideList(PowerPointDocument owner) : IReadOnlyList<PowerPointSlide> {
            public int Count {
                get {
                    owner.ThrowIfDisposed();
                    return owner._sdkDocument == null && owner._sdkDeferred ? owner._deferredSlideCount : owner.LoadedSlides.Count;
                }
            }

            public PowerPointSlide this[int index] {
                get {
                    owner.ThrowIfDisposed();
                    return (owner.GetPackageSlides() ?? owner.LoadedSlides)[index];
                }
            }

            public IEnumerator<PowerPointSlide> GetEnumerator() {
                owner.ThrowIfDisposed();
                // A copy, so the slides read while enumerating may change the list.
                return (owner.GetPackageSlides()?.ToList() ?? owner.LoadedSlides).GetEnumerator();
            }

            System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
        }

        /// <summary>
        ///     Built-in package properties for the presentation.
        /// </summary>
        public PowerPointBuiltinDocumentProperties Properties => _properties ??= new PowerPointBuiltinDocumentProperties(SdkDocument);

        /// <summary>
        ///     Extended application properties for the presentation.
        /// </summary>
        public PowerPointApplicationProperties ApplicationProperties => _applicationProperties ??= new PowerPointApplicationProperties(SdkDocument);

        /// <summary>
        ///     Slide size information for the presentation.
        /// </summary>
        public PowerPointSlideSize SlideSize {
            get {
                ThrowIfDisposed();
                return _slideSize ??= _sdkDocument == null && _sdkDeferred ? new PowerPointSlideSize(this) : new PowerPointSlideSize(SdkPresentationPart);
            }
        }

        /// <summary>
        ///     Replaces literal text on every slide.
        /// </summary>
        /// <param name="text">Text to find.</param>
        /// <param name="replacement">Replacement text; <c>null</c> removes each match.</param>
        /// <param name="options">Comparison settings and scope. <c>null</c> matches ordinally, ignoring case, inside or across words, in textboxes and tables but not notes.</param>
        /// <returns>The number of replaced matches.</returns>
        /// <remarks>Matches in textboxes and tables can span adjacent runs within one paragraph. See <see cref="PowerPointSlide.ReplaceAll"/> for scope and formatting.</remarks>
        public int ReplaceAll(string text, string replacement, PowerPointTextSearchOptions? options = null) {
            ThrowIfDisposed();
            OfficeTextSearch.ValidateSearchText(text, nameof(text));

            if (TryReplaceAllWithoutSdk(text, replacement, options, out int count)) return count;
            foreach (PowerPointSlide slide in Slides) {
                count += slide.ReplaceAll(text, replacement, options);
            }
            return count;
        }

        private void ThrowIfDisposed() {
            if (_disposed) {
                throw new ObjectDisposedException(nameof(PowerPointDocument));
            }
        }
    }
}
