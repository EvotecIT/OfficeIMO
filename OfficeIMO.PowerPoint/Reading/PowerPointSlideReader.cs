using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Reads the slides of a PPTX, PPTM, POTX or PPSX package one at a time without building the presentation
    ///     model or loading the Open XML SDK. Each slide part is read forward when its turn comes and its XML is
    ///     released once read, so the first slide is available before the rest of the deck is read. Relationship
    ///     metadata of the slides read so far stays cached until the reader is disposed.
    /// </summary>
    /// <remarks>
    ///     <para>
    ///         Text is read as authored on each slide: placeholders that are empty on the slide do not pick up layout
    ///         or master prompt text, and list markers come from the paragraph's own properties. Text inside SmartArt
    ///         diagrams, charts and markup-compatibility fallbacks is not read.
    ///     </para>
    ///     <para>
    ///         Relationships are read per slide when the slide is read, not for the whole package when it opens.
    ///         The package is opened without the repair step a model load applies, so a package whose relationship
    ///         targets are not valid URIs fails when the part that holds them is read; <see cref="PowerPointDocument.Load(string, PowerPointLoadOptions?)"/>
    ///         reads such packages.
    ///         The reader keeps the package open until it is disposed. A file is opened for shared reading; a caller
    ///         stream stays open and owned by the caller.
    ///     </para>
    /// </remarks>
    public sealed partial class PowerPointSlideReader : IDisposable {
        private readonly OfficeOpcPackage _package;
        private readonly Stream? _ownedStream;
        private readonly IReadOnlyList<SlideEntry> _slides;
        private readonly Dictionary<string, int> _slideNumbers;
        private readonly long _maximumCharacters;
        private bool _disposed;

        private const string NotesSlideContentType = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";
        private readonly bool _hasNotesParts;

        private PowerPointSlideReader(OfficeOpcPackage package, Stream? ownedStream, IReadOnlyList<SlideEntry> slides,
            long maximumCharacters) {
            _package = package;
            _ownedStream = ownedStream;
            _slides = slides;
            _maximumCharacters = maximumCharacters;
            // A notes slide has the notes-slide content type, or the package is not one the model opens; without such
            // a part no slide has notes, and the slides' relationships need not be read to find them.
            _hasNotesParts = package.DeclaresContentType(NotesSlideContentType);
            // The first slide-list entry of each part numbers it, as slide hyperlinks resolve in the model.
            // Slides are keyed by part identity, so a hyperlink that spells the part with percent-encoding still finds it.
            _slideNumbers = new Dictionary<string, int>(OfficeAsciiIgnoreCaseComparer.Instance);
            for (int index = 0; index < slides.Count; index++) {
                string key = OfficeOpcPackage.GetPartIdentity(slides[index].PartName);
                if (!_slideNumbers.ContainsKey(key)) _slideNumbers.Add(key, index + 1);
            }
        }

        /// <summary>Number of slides in presentation order.</summary>
        public int SlideCount {
            get {
                ThrowIfDisposed();
                return _slides.Count;
            }
        }

        /// <summary>Opens a presentation file for streaming slide reads with the default input limits.</summary>
        /// <param name="path">Path to an Open XML presentation.</param>
        public static PowerPointSlideReader Open(string path) => Open(path, null, CancellationToken.None);

        /// <summary>Opens a presentation stream for streaming slide reads with the default input limits.</summary>
        /// <param name="stream">Readable presentation stream; see <see cref="Open(Stream, PowerPointLoadOptions?, CancellationToken)"/>.</param>
        public static PowerPointSlideReader Open(Stream stream) => Open(stream, null, CancellationToken.None);

        /// <summary>Opens a presentation file for streaming slide reads.</summary>
        /// <param name="path">Path to an Open XML presentation.</param>
        /// <param name="options">
        ///     Input limits: <see cref="PowerPointLoadOptions.MaxInputBytes"/>,
        ///     <see cref="DocumentLoadOptions.PackageSecurity"/> and
        ///     <see cref="OfficeOpenXmlLoadSettings.MaxCharactersInPart"/> apply as they do to a model load.
        ///     Passwords and binary PPT input are not supported.
        /// </param>
        /// <param name="cancellationToken">Cancellation for opening the package.</param>
        public static PowerPointSlideReader Open(string path, PowerPointLoadOptions? options,
            CancellationToken cancellationToken) {
            ArgumentNullException.ThrowIfNull(path);
            if (!File.Exists(path)) throw new FileNotFoundException($"File '{path}' doesn't exist.", path);
            var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read | FileShare.Delete);
            try {
                return Open(stream, ownsStream: true, options, cancellationToken);
            } catch {
                stream.Dispose();
                throw;
            }
        }

        /// <summary>Opens a presentation stream for streaming slide reads.</summary>
        /// <param name="stream">
        ///     Readable presentation stream. A seekable stream is read in place and must stay open and unchanged while
        ///     the reader is used; a forward-only stream is first copied to memory within
        ///     <see cref="PowerPointLoadOptions.MaxInputBytes"/>.
        /// </param>
        /// <param name="options">Input limits, as for <see cref="Open(string, PowerPointLoadOptions?, CancellationToken)"/>.</param>
        /// <param name="cancellationToken">Cancellation for opening the package.</param>
        public static PowerPointSlideReader Open(Stream stream, PowerPointLoadOptions? options,
            CancellationToken cancellationToken) {
            ArgumentNullException.ThrowIfNull(stream);
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
            if (stream.CanSeek) return Open(stream, ownsStream: false, options, cancellationToken);
            PowerPointLoadOptions resolved = options ?? new PowerPointLoadOptions();
            var buffered = new MemoryStream(OfficeStreamReader.ReadRemainingBytes(stream, cancellationToken, resolved.MaxInputBytes), writable: false);
            try {
                return Open(buffered, ownsStream: true, resolved, cancellationToken);
            } catch {
                buffered.Dispose();
                throw;
            }
        }

        private static PowerPointSlideReader Open(Stream stream, bool ownsStream, PowerPointLoadOptions? options,
            CancellationToken cancellationToken) {
            PowerPointLoadOptions resolved = options ?? new PowerPointLoadOptions();
            if (resolved.Password != null) {
                throw new NotSupportedException("Streaming slide reads do not open encrypted presentations. Use PowerPointDocument.Load with a password.");
            }
            cancellationToken.ThrowIfCancellationRequested();
            // Like a model load, the whole stream is the package whatever its current position.
            stream.Position = 0;
            long length = stream.Length;
            if (resolved.MaxInputBytes.HasValue && length > resolved.MaxInputBytes.Value) {
                throw new InvalidDataException($"Stream exceeds the configured maximum size ({resolved.MaxInputBytes.Value} bytes).");
            }
            if (!HasZipSignature(stream)) {
                throw new NotSupportedException("Streaming slide reads require an Open XML presentation package. Load binary PPT files with PowerPointDocument.Load.");
            }
            if (resolved.PackageSecurity != null) OfficePackageSecurityInspector.Validate(stream, resolved.PackageSecurity);
            cancellationToken.ThrowIfCancellationRequested();

            long maximumCharacters = resolved.OpenSettings?.MaxCharactersInPart > 0
                ? Math.Min(PowerPointXmlReader.MaximumPackageXmlCharacters, resolved.OpenSettings.MaxCharactersInPart)
                : PowerPointXmlReader.MaximumPackageXmlCharacters;
            OfficeOpcPackage package = OfficeOpcPackage.Open(stream, leaveOpen: true);
            try {
                string presentationPart = GetPresentationPart(package);
                IReadOnlyList<SlideEntry> slides = ReadSlideList(package, presentationPart, maximumCharacters, cancellationToken);
                return new PowerPointSlideReader(package, ownsStream ? stream : null, slides, maximumCharacters);
            } catch {
                package.Dispose();
                throw;
            }
        }

        /// <summary>
        ///     Reads slides in presentation order. Each slide is read when the enumerator reaches it; nothing is kept
        ///     after it is returned.
        /// </summary>
        /// <param name="includeNotes">Reads each slide's speaker notes part. Default: true.</param>
        /// <param name="cancellationToken">Cancellation checked between slides and shapes.</param>
        public IEnumerable<PowerPointSlideContent> ReadSlides(bool includeNotes = true,
            CancellationToken cancellationToken = default) {
            ThrowIfDisposed();
            return ReadSlidesCore(includeNotes, cancellationToken);
        }

        private IEnumerable<PowerPointSlideContent> ReadSlidesCore(bool includeNotes, CancellationToken cancellationToken) {
            for (int index = 0; index < _slides.Count; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                ThrowIfDisposed();
                yield return ReadSlide(index, includeNotes, cancellationToken);
            }
        }

        /// <summary>
        ///     Extracts slide-aligned Markdown chunks while the slides are read, so each chunk is returned as soon as
        ///     its slide has been read. The chunks match
        ///     <see cref="PowerPointExtractionExtensions.ExtractMarkdownChunks(PowerPointDocument, PowerPointExtractionExtensions.PowerPointExtractOptions?, PowerPointExtractChunkingOptions?, string?, CancellationToken)"/>
        ///     for the same presentation.
        /// </summary>
        /// <param name="extract">Notes, tables, hidden shapes and Markdown options.</param>
        /// <param name="chunking">Chunk size options.</param>
        /// <param name="sourcePath">Optional source path used in chunk identifiers and locations.</param>
        /// <param name="cancellationToken">Cancellation checked between slides and shapes.</param>
        public IEnumerable<PowerPointExtractChunk> ExtractMarkdownChunks(
            PowerPointExtractionExtensions.PowerPointExtractOptions? extract = null,
            PowerPointExtractChunkingOptions? chunking = null, string? sourcePath = null,
            CancellationToken cancellationToken = default) {
            ThrowIfDisposed();
            return PowerPointExtractionExtensions.ExtractMarkdownChunks(this, extract, chunking, sourcePath, cancellationToken);
        }

        /// <summary>Core document properties of the package.</summary>
        internal (string? Title, string? Creator, string? Subject, string? Keywords) ReadProperties() {
            ThrowIfDisposed();
            OfficeOpcCoreProperties properties = _package.ReadCoreProperties();
            return (properties.Title, properties.Creator, properties.Subject, properties.Keywords);
        }

        /// <summary>Closes the package and any stream the reader opened.</summary>
        public void Dispose() {
            if (_disposed) return;
            _disposed = true;
            try {
                _package.Dispose();
            } finally {
                _ownedStream?.Dispose();
            }
        }

        private void ThrowIfDisposed() {
            if (_disposed) throw new ObjectDisposedException(nameof(PowerPointSlideReader));
        }

        private static bool HasZipSignature(Stream stream) {
            var signature = new byte[2];
            int read = 0;
            while (read < signature.Length) {
                int count = stream.Read(signature, read, signature.Length - read);
                if (count == 0) break;
                read += count;
            }
            stream.Position = 0;
            return read == 2 && signature[0] == 0x50 && signature[1] == 0x4B;
        }

        private readonly struct SlideEntry {
            internal SlideEntry(string partName, string? listShowValue) {
                PartName = partName;
                ListShowValue = listShowValue;
            }

            internal string PartName { get; }

            // The legacy show attribute some producers write on the slide list entry instead of the slide.
            internal string? ListShowValue { get; }
        }
    }
}
