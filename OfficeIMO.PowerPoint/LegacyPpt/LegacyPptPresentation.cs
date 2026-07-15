using OfficeIMO.Drawing.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Diagnostics;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt {
    /// <summary>
    /// Represents the dependency-free, normalized contents of a PowerPoint 97-2003 binary presentation.
    /// </summary>
    public sealed class LegacyPptPresentation {
        private const ushort RecordDocument = 0x03E8;
        private const ushort RecordDocumentAtom = 0x03E9;
        private const ushort RecordSlide = 0x03EE;
        private const ushort RecordSlideAtom = 0x03EF;
        private const ushort RecordSlidePersistAtom = 0x03F3;
        private const ushort RecordMainMaster = 0x03F8;
        private const ushort RecordSlideShowSlideInfoAtom = 0x03F9;
        private const ushort RecordDrawing = 0x040C;
        private const ushort RecordPlaceholder = 0x0BC3;
        private const ushort RecordTextChars = 0x0FA0;
        private const ushort RecordStyleTextPropAtom = 0x0FA1;
        private const ushort RecordTextBytes = 0x0FA8;
        private const ushort RecordSlideListWithText = 0x0FF0;
        private const ushort OfficeArtSpContainer = 0xF004;
        private const ushort OfficeArtFsp = 0xF00A;
        private const ushort OfficeArtFopt = 0xF00B;
        private const ushort OfficeArtClientTextbox = 0xF00D;
        private const ushort OfficeArtChildAnchor = 0xF00F;
        private const ushort OfficeArtClientAnchor = 0xF010;

        private readonly List<LegacyPptSlide> _slides = new();
        private readonly List<LegacyPptMaster> _masters = new();
        private readonly List<LegacyPptImportDiagnostic> _diagnostics = new();

        private LegacyPptPresentation() { }

        internal LegacyPptPackage Package { get; private set; } = null!;

        /// <summary>Gets the slide width in PowerPoint master units (576 units per inch).</summary>
        public int SlideWidth { get; private set; } = 7200;

        /// <summary>Gets the slide height in PowerPoint master units (576 units per inch).</summary>
        public int SlideHeight { get; private set; } = 5400;

        /// <summary>Gets the decoded slides in display order.</summary>
        public IReadOnlyList<LegacyPptSlide> Slides => _slides;

        /// <summary>Gets decoded main masters and title masters in document order.</summary>
        public IReadOnlyList<LegacyPptMaster> Masters => _masters;

        /// <summary>Gets import diagnostics, including preserve-only content.</summary>
        public IReadOnlyList<LegacyPptImportDiagnostic> Diagnostics => _diagnostics;

        /// <summary>Loads a PowerPoint 97-2003 binary presentation from a path.</summary>
        public static LegacyPptPresentation Load(string path, LegacyPptImportOptions? options = null) {
            if (path == null) throw new ArgumentNullException(nameof(path));
            return Load(File.ReadAllBytes(path), options);
        }

        /// <summary>Loads a PowerPoint 97-2003 binary presentation from a stream.</summary>
        public static LegacyPptPresentation Load(Stream stream, LegacyPptImportOptions? options = null) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
            return Load(OfficeStreamReader.ReadAllBytes(stream), options);
        }

        /// <summary>Loads a PowerPoint 97-2003 binary presentation from bytes.</summary>
        public static LegacyPptPresentation Load(byte[] bytes, LegacyPptImportOptions? options = null) {
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));
            options ??= new LegacyPptImportOptions();
            LegacyPptPackage package = LegacyPptPackage.Read(bytes, options);
            var presentation = new LegacyPptPresentation { Package = package };
            presentation.AddCompoundFeatureDiagnostics(package.CompoundFile, options);
            presentation.Parse(package, options);
            return presentation;
        }

        /// <summary>Creates a compact import inventory.</summary>
        public LegacyPptImportReport CreateImportReport() => new LegacyPptImportReport(this);

        private void Parse(LegacyPptPackage package, LegacyPptImportOptions options) {
            byte[] documentStream = package.DocumentStream;
            IReadOnlyDictionary<uint, uint> persistOffsets = package.PersistObjectOffsets;
            uint documentPersistId = package.DocumentPersistId;
            if (!persistOffsets.TryGetValue(documentPersistId, out uint documentOffset)) {
                throw new InvalidDataException($"The document persist object {documentPersistId} is missing.");
            }

            LegacyPptRecord document = LegacyPptRecordReader.ReadSingle(documentStream,
                ToBoundedOffset(documentOffset, documentStream.Length, "document persist object"), options);
            if (document.Type != RecordDocument) {
                throw new InvalidDataException($"Persist object {documentPersistId} is record 0x{document.Type:X4}, not DocumentContainer.");
            }

            LegacyPptRecord? documentAtom = document.Children.FirstOrDefault(record => record.Type == RecordDocumentAtom);
            if (documentAtom != null && documentAtom.PayloadLength >= 8) {
                int width = documentAtom.ReadInt32(0);
                int height = documentAtom.ReadInt32(4);
                if (width > 0 && height > 0) {
                    SlideWidth = width;
                    SlideHeight = height;
                }
            }

            ParseMasters(document, documentStream, persistOffsets, options);

            LegacyPptRecord? slideList = document.Children.FirstOrDefault(record =>
                record.Type == RecordSlideListWithText && record.Instance == 0);
            if (slideList == null) {
                AddDiagnostic("PPT-SLIDES-MISSING", LegacyPptDiagnosticSeverity.Warning,
                    "The document has no slide list.", document.Offset);
                return;
            }

            int slideIndex = 0;
            foreach (LegacyPptRecord slidePersist in slideList.Children.Where(record => record.Type == RecordSlidePersistAtom)) {
                if (slidePersist.PayloadLength < 16) {
                    AddDiagnostic("PPT-SLIDE-PERSIST-TRUNCATED", LegacyPptDiagnosticSeverity.Warning,
                        "A slide directory entry is truncated and was skipped.", slidePersist.Offset);
                    continue;
                }
                uint persistId = slidePersist.ReadUInt32(0);
                uint slideId = slidePersist.ReadUInt32(12);
                if (!persistOffsets.TryGetValue(persistId, out uint slideOffset)) {
                    AddDiagnostic("PPT-SLIDE-PERSIST-MISSING", LegacyPptDiagnosticSeverity.Warning,
                        $"Slide {slideId} references missing persist object {persistId}.", slidePersist.Offset);
                    continue;
                }

                LegacyPptRecord slideRecord = LegacyPptRecordReader.ReadSingle(documentStream,
                    ToBoundedOffset(slideOffset, documentStream.Length, "slide persist object"), options);
                if (slideRecord.Type != RecordSlide) {
                    AddDiagnostic("PPT-SLIDE-TYPE", LegacyPptDiagnosticSeverity.Warning,
                        $"Slide {slideId} points to record 0x{slideRecord.Type:X4}; the slide was skipped.", slideRecord.Offset);
                    continue;
                }

                var slide = new LegacyPptSlide(slideId, persistId) { Name = $"Slide {++slideIndex}" };
                ParseSlide(slideRecord, slide, options);
                TryReadNotes(slideRecord, slide, documentStream, persistOffsets, options);
                _slides.Add(slide);
            }
        }

        private void ParseSlide(LegacyPptRecord slideRecord, LegacyPptSlide slide, LegacyPptImportOptions options) {
            LegacyPptRecord? slideAtom = slideRecord.Children.FirstOrDefault(record => record.Type == RecordSlideAtom);
            if (slideAtom != null && slideAtom.PayloadLength >= 16) {
                slide.MasterId = slideAtom.ReadUInt32(12);
            }
            LegacyPptRecord? slideShowInfo = slideRecord.Children.FirstOrDefault(record =>
                record.Type == RecordSlideShowSlideInfoAtom);
            if (slideShowInfo != null && slideShowInfo.PayloadLength >= 11) {
                slide.Hidden = (slideShowInfo.ReadByte(10) & 0x04) != 0;
            }
            ParseShapes(slideRecord, slide.AddShape, "slide", options);
        }

        private void ParseMasters(LegacyPptRecord document, byte[] documentStream,
            IReadOnlyDictionary<uint, uint> persistOffsets, LegacyPptImportOptions options) {
            LegacyPptRecord? masterList = document.Children.FirstOrDefault(record =>
                record.Type == RecordSlideListWithText && record.Instance == 1);
            if (masterList == null) return;

            foreach (LegacyPptRecord masterPersist in masterList.Children.Where(record =>
                         record.Type == RecordSlidePersistAtom)) {
                if (masterPersist.PayloadLength < 20) {
                    AddDiagnostic("PPT-MASTER-PERSIST-TRUNCATED", LegacyPptDiagnosticSeverity.Warning,
                        "A master directory entry is truncated and was skipped.", masterPersist.Offset);
                    continue;
                }
                uint persistId = masterPersist.ReadUInt32(0);
                uint masterId = masterPersist.ReadUInt32(12);
                if (!persistOffsets.TryGetValue(persistId, out uint masterOffset)) {
                    AddDiagnostic("PPT-MASTER-PERSIST-MISSING", LegacyPptDiagnosticSeverity.Warning,
                        $"Master 0x{masterId:X8} references missing persist object {persistId}.", masterPersist.Offset);
                    continue;
                }

                LegacyPptRecord masterRecord = LegacyPptRecordReader.ReadSingle(documentStream,
                    ToBoundedOffset(masterOffset, documentStream.Length, "master persist object"), options);
                bool isMainMaster = masterRecord.Type == RecordMainMaster;
                if (!isMainMaster && masterRecord.Type != RecordSlide) {
                    AddDiagnostic("PPT-MASTER-TYPE", LegacyPptDiagnosticSeverity.Warning,
                        $"Master 0x{masterId:X8} points to record 0x{masterRecord.Type:X4}; it was skipped.",
                        masterRecord.Offset);
                    continue;
                }

                LegacyPptRecord? slideAtom = masterRecord.Children.FirstOrDefault(record =>
                    record.Type == RecordSlideAtom);
                uint parentMasterId = slideAtom != null && slideAtom.PayloadLength >= 16
                    ? slideAtom.ReadUInt32(12)
                    : 0U;
                var master = new LegacyPptMaster(masterId, persistId, isMainMaster, parentMasterId);
                ParseShapes(masterRecord, master.AddShape, isMainMaster ? "main master" : "title master", options);
                _masters.Add(master);
            }
        }

        private void ParseShapes(LegacyPptRecord ownerRecord, Action<LegacyPptShape> addShape,
            string ownerDescription, LegacyPptImportOptions options) {
            LegacyPptRecord? drawing = ownerRecord.Children.FirstOrDefault(record => record.Type == RecordDrawing);
            if (drawing == null) return;

            LegacyPptRecord? primaryShapeGroup = drawing.DescendantsAndSelf()
                .FirstOrDefault(record => record.Type == 0xF003);
            if (primaryShapeGroup == null) return;

            if (primaryShapeGroup.Children.Any(record => record.Type == 0xF003)) {
                AddDiagnostic("PPT-GROUP-UNSUPPORTED", LegacyPptDiagnosticSeverity.Warning,
                    $"Grouped OfficeArt shapes on a {ownerDescription} are not projected because their nested coordinate system is not yet supported.",
                    primaryShapeGroup.Offset);
            }

            foreach (LegacyPptRecord shapeContainer in primaryShapeGroup.Children
                         .Where(record => record.Type == OfficeArtSpContainer)) {
                LegacyPptRecord? fsp = shapeContainer.Children.FirstOrDefault(record => record.Type == OfficeArtFsp);
                LegacyPptRecord? anchor = shapeContainer.Children.FirstOrDefault(record =>
                    record.Type == OfficeArtClientAnchor || record.Type == OfficeArtChildAnchor);
                if (fsp == null || anchor == null || fsp.PayloadLength < 8) continue;

                ushort shapeType = fsp.Instance;
                uint shapeId = fsp.ReadUInt32(0);
                LegacyPptBounds bounds;
                try {
                    bounds = ReadBounds(anchor);
                } catch (InvalidDataException) {
                    AddDiagnostic("PPT-SHAPE-ANCHOR", LegacyPptDiagnosticSeverity.Warning,
                        "A shape has an unsupported or truncated anchor and was skipped.", anchor.Offset);
                    continue;
                }

                LegacyPptRecord? textbox = shapeContainer.Children.FirstOrDefault(record =>
                    record.Type == OfficeArtClientTextbox);
                string text = textbox == null ? string.Empty : ReadText(textbox);
                LegacyPptPlaceholderKind placeholder = ReadPlaceholder(shapeContainer);
                LegacyPptShapeKind kind = ClassifyShape(shapeType, textbox != null || text.Length > 0);
                addShape(new LegacyPptShape(kind, shapeType, shapeId, shapeContainer.Offset,
                    bounds, text, placeholder));

                if (textbox != null && textbox.DescendantsAndSelf()
                        .Any(record => record.Type == RecordStyleTextPropAtom)) {
                    AddDiagnostic("PPT-TEXT-FORMATTING-FLATTENED", LegacyPptDiagnosticSeverity.Warning,
                        "Rich text and paragraph formatting was flattened to plain text.", textbox.Offset);
                }
                if (shapeContainer.Children.Any(record => record.Type == OfficeArtFopt)) {
                    AddDiagnostic("PPT-SHAPE-STYLE-FLATTENED", LegacyPptDiagnosticSeverity.Warning,
                        "OfficeArt fill, outline, or text-box styling was not projected.", shapeContainer.Offset);
                }

                if (kind == LegacyPptShapeKind.Unsupported && options.ReportUnsupportedContent) {
                    AddDiagnostic("PPT-SHAPE-UNSUPPORTED", LegacyPptDiagnosticSeverity.Warning,
                        $"OfficeArt shape type {shapeType} on a {ownerDescription} is not projected to the editable PowerPoint model.", shapeContainer.Offset);
                }
            }
        }

        private void TryReadNotes(LegacyPptRecord slideRecord, LegacyPptSlide slide, byte[] documentStream,
            IReadOnlyDictionary<uint, uint> persistOffsets, LegacyPptImportOptions options) {
            LegacyPptRecord? slideAtom = slideRecord.Children.FirstOrDefault(record => record.Type == RecordSlideAtom);
            if (slideAtom == null || slideAtom.PayloadLength < 20) return;
            uint notesPersistId = slideAtom.ReadUInt32(16);
            if (notesPersistId == 0 || !persistOffsets.TryGetValue(notesPersistId, out uint notesOffset)) return;

            try {
                LegacyPptRecord notes = LegacyPptRecordReader.ReadSingle(documentStream,
                    ToBoundedOffset(notesOffset, documentStream.Length, "notes persist object"), options);
                string notesText = string.Join("\n", notes.DescendantsAndSelf()
                    .Where(record => record.Type == OfficeArtClientTextbox)
                    .Select(ReadText)
                    .Where(text => !string.IsNullOrWhiteSpace(text)));
                slide.NotesText = notesText;
            } catch (InvalidDataException exception) {
                AddDiagnostic("PPT-NOTES-READ", LegacyPptDiagnosticSeverity.Warning,
                    $"Speaker notes could not be decoded: {exception.Message}", notesOffset);
            }
        }

        private static LegacyPptBounds ReadBounds(LegacyPptRecord anchor) {
            if (anchor.PayloadLength == 8) {
                int top = anchor.ReadInt16(0);
                int left = anchor.ReadInt16(2);
                int right = anchor.ReadInt16(4);
                int bottom = anchor.ReadInt16(6);
                return CreateBounds(left, top, right, bottom);
            }
            if (anchor.PayloadLength >= 16) {
                int top = anchor.ReadInt32(0);
                int left = anchor.ReadInt32(4);
                int right = anchor.ReadInt32(8);
                int bottom = anchor.ReadInt32(12);
                return CreateBounds(left, top, right, bottom);
            }
            throw new InvalidDataException("The OfficeArt anchor is too short.");
        }

        private static LegacyPptBounds CreateBounds(int left, int top, int right, int bottom) =>
            new LegacyPptBounds(left, top, Math.Max(0, right - left), Math.Max(0, bottom - top));

        private static string ReadText(LegacyPptRecord textbox) {
            LegacyPptRecord? textRecord = textbox.DescendantsAndSelf().FirstOrDefault(record =>
                record.Type == RecordTextChars || record.Type == RecordTextBytes);
            if (textRecord == null) return string.Empty;
            string text = textRecord.Type == RecordTextChars
                ? textRecord.ReadUtf16Text()
                : textRecord.ReadLowByteUnicodeText();
            return text.TrimEnd('\0').Replace("\r", "\n").TrimEnd('\n');
        }

        private static LegacyPptPlaceholderKind ReadPlaceholder(LegacyPptRecord shapeContainer) {
            LegacyPptRecord? placeholder = shapeContainer.DescendantsAndSelf()
                .FirstOrDefault(record => record.Type == RecordPlaceholder);
            if (placeholder == null || placeholder.PayloadLength < 5) return LegacyPptPlaceholderKind.None;
            byte value = placeholder.ReadByte(4);
            return Enum.IsDefined(typeof(LegacyPptPlaceholderKind), value)
                ? (LegacyPptPlaceholderKind)value
                : LegacyPptPlaceholderKind.None;
        }

        private static LegacyPptShapeKind ClassifyShape(ushort shapeType, bool hasText) {
            if (hasText || shapeType == 202) return LegacyPptShapeKind.TextBox;
            if (shapeType == 1) return LegacyPptShapeKind.Rectangle;
            if (shapeType == 3) return LegacyPptShapeKind.Ellipse;
            if (shapeType == 20) return LegacyPptShapeKind.Line;
            return LegacyPptShapeKind.Unsupported;
        }

        private void AddCompoundFeatureDiagnostics(OfficeCompoundFile compound, LegacyPptImportOptions options) {
            if (!options.ReportUnsupportedContent) return;
            if (compound.Streams.Keys.Any(name => name.IndexOf("VBA", StringComparison.OrdinalIgnoreCase) >= 0)) {
                AddDiagnostic("PPT-VBA-PRESERVE-ONLY", LegacyPptDiagnosticSeverity.Warning,
                    "The compound file contains a VBA project that is not projected into the editable model.", null);
            }
            if (compound.Streams.Keys.Any(name => name.IndexOf("ObjectPool", StringComparison.OrdinalIgnoreCase) >= 0)) {
                AddDiagnostic("PPT-OLE-PRESERVE-ONLY", LegacyPptDiagnosticSeverity.Warning,
                    "The compound file contains embedded OLE objects that are not projected into the editable model.", null);
            }
        }

        private static int ToBoundedOffset(uint offset, int length, string description) {
            if (offset > int.MaxValue || offset > unchecked((uint)Math.Max(0, length - 8))) {
                throw new InvalidDataException($"The {description} offset 0x{offset:X} is outside the PowerPoint Document stream.");
            }
            return unchecked((int)offset);
        }

        private void AddDiagnostic(string code, LegacyPptDiagnosticSeverity severity, string message, long? offset) {
            _diagnostics.Add(new LegacyPptImportDiagnostic(code, message, severity, offset));
        }
    }
}
