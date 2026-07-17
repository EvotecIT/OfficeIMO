using OfficeIMO.Drawing.Internal;
using OfficeIMO.Drawing.Binary;
using OfficeIMO.PowerPoint.LegacyPpt.Diagnostics;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt {
    /// <summary>
    /// Represents the dependency-free, normalized contents of a PowerPoint 97-2003 binary presentation.
    /// </summary>
    public sealed partial class LegacyPptPresentation {
        private const ushort RecordDocument = 0x03E8;
        private const ushort RecordDocumentAtom = 0x03E9;
        private const ushort RecordSlide = 0x03EE;
        private const ushort RecordSlideAtom = 0x03EF;
        private const ushort RecordSlidePersistAtom = 0x03F3;
        private const ushort RecordMainMaster = 0x03F8;
        private const ushort RecordSlideShowSlideInfoAtom = 0x03F9;
        private const ushort RecordColorSchemeAtom = 0x07F0;
        private const ushort RecordFontCollection = 0x07D5;
        private const ushort RecordSoundCollection = 0x07E4;
        private const ushort RecordSoundCollectionAtom = 0x07E5;
        private const ushort RecordSound = 0x07E6;
        private const ushort RecordSoundDataBlob = 0x07E7;
        private const ushort RecordDrawing = 0x040C;
        private const ushort RecordPlaceholder = 0x0BC3;
        private const ushort RecordTextHeaderAtom = 0x0F9F;
        private const ushort RecordTextChars = 0x0FA0;
        private const ushort RecordStyleTextPropAtom = 0x0FA1;
        private const ushort RecordTextMasterStyleAtom = 0x0FA3;
        private const ushort RecordTextRulerAtom = 0x0FA6;
        private const ushort RecordTextBytes = 0x0FA8;
        private const ushort RecordTextSpecialInfoAtom = 0x0FAA;
        private const ushort RecordFontEntityAtom = 0x0FB7;
        private const ushort RecordFontEmbedDataBlob = 0x0FB8;
        private const ushort RecordSlideListWithText = 0x0FF0;
        private const ushort OfficeArtSpContainer = 0xF004;
        private const ushort OfficeArtSpgrContainer = 0xF003;
        private const ushort OfficeArtSolverContainer = 0xF005;
        private const ushort OfficeArtFConnectorRule = 0xF012;
        private const ushort OfficeArtBStoreContainer = 0xF001;
        private const ushort OfficeArtFbse = 0xF007;
        private const ushort OfficeArtFspgr = 0xF009;
        private const ushort OfficeArtFsp = 0xF00A;
        private const ushort OfficeArtFopt = 0xF00B;
        private const ushort OfficeArtTertiaryFopt = 0xF122;
        private const ushort OfficeArtClientTextbox = 0xF00D;
        private const ushort OfficeArtChildAnchor = 0xF00F;
        private const ushort OfficeArtClientAnchor = 0xF010;

        private readonly List<LegacyPptSlide> _slides = new();
        private readonly List<LegacyPptMaster> _masters = new();
        private readonly List<OfficeArtBlipStoreEntry> _blipStoreEntries = new();
        private readonly List<LegacyPptPictureBullet> _pictureBullets = new();
        private readonly Dictionary<ushort, LegacyPptPictureBullet>
            _pictureBulletsByIndex = new();
        private readonly List<LegacyPptFont> _fonts = new();
        private readonly Dictionary<ushort, LegacyPptFont> _fontsByIndex = new();
        private readonly List<LegacyPptSound> _sounds = new();
        private readonly Dictionary<uint, LegacyPptSound> _soundsById = new();
        private readonly List<LegacyPptImportDiagnostic> _diagnostics = new();
        private LegacyPptRecordTraversalBudget _recordBudget = null!;
        private LegacyPptDecodedStorageBudget _decodedStorageBudget = null!;

        private LegacyPptPresentation() { }

        internal LegacyPptPackage Package { get; private set; } = null!;

        /// <summary>Gets the slide width in PowerPoint master units (576 units per inch).</summary>
        public int SlideWidth { get; private set; } = 7200;

        /// <summary>Gets the slide height in PowerPoint master units (576 units per inch).</summary>
        public int SlideHeight { get; private set; } = 5400;

        /// <summary>Gets the complete binary document settings when the DocumentAtom is valid.</summary>
        public LegacyPptDocumentSettings? DocumentSettings { get; private set; }

        /// <summary>Gets the decoded slides in display order.</summary>
        public IReadOnlyList<LegacyPptSlide> Slides => _slides;

        /// <summary>Gets decoded main masters and title masters in document order.</summary>
        public IReadOnlyList<LegacyPptMaster> Masters => _masters;

        /// <summary>Gets the document-level OfficeArt BLIP store in one-based reference order.</summary>
        public IReadOnlyList<OfficeArtBlipStoreEntry> BlipStoreEntries => _blipStoreEntries;

        /// <summary>Gets PPT9 document-level picture bullets by their sparse zero-based indexes.</summary>
        public IReadOnlyList<LegacyPptPictureBullet> PictureBullets =>
            _pictureBullets;

        /// <summary>Gets document fonts by their binary PowerPoint font index.</summary>
        public IReadOnlyList<LegacyPptFont> Fonts => _fonts;

        /// <summary>Gets the document-level sounds referenced by transitions and interactive actions.</summary>
        public IReadOnlyList<LegacyPptSound> Sounds => _sounds;

        /// <summary>Gets the sound identifier seed stored by the document, when valid.</summary>
        public uint? SoundIdSeed { get; private set; }

        /// <summary>Gets the complete VBA project storage referenced by the presentation, when present.</summary>
        public LegacyPptVbaProject? VbaProject { get; private set; }

        /// <summary>Gets import diagnostics, including preserve-only content.</summary>
        public IReadOnlyList<LegacyPptImportDiagnostic> Diagnostics => _diagnostics;

        /// <summary>Gets whether the source package was password-encrypted with binary RC4 CryptoAPI encryption.</summary>
        public bool WasEncryptedSource => Package.WasEncryptedSource;

        /// <summary>Gets the source RC4 key size in bits, or <see langword="null"/> for an unencrypted source.</summary>
        public int? EncryptionKeySizeBits => Package.EncryptionKeySizeBits;

        /// <summary>
        /// Gets whether the encrypted source protected its document-property streams,
        /// or <see langword="null"/> for an unencrypted source.
        /// </summary>
        public bool? EncryptedDocumentProperties =>
            Package.EncryptedDocumentProperties;

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
            var recordBudget = new LegacyPptRecordTraversalBudget(
                options.MaxRecordCount);
            var decodedStorageBudget = new LegacyPptDecodedStorageBudget(
                options.MaxDecodedStorageBytes);
            LegacyPptPackage package = LegacyPptPackage.Read(bytes, options,
                recordBudget);
            var presentation = new LegacyPptPresentation {
                Package = package,
                _recordBudget = recordBudget,
                _decodedStorageBudget = decodedStorageBudget
            };
            if (package.WasEncryptedSource) {
                presentation.AddDiagnostic("PPT-ENCRYPTION-DECRYPTED",
                    LegacyPptDiagnosticSeverity.Information,
                    $"The RC4 CryptoAPI encrypted presentation was decrypted with its {package.EncryptionKeySizeBits}-bit key descriptor.",
                    null);
            }
            presentation.AddCompoundFeatureDiagnostics(package.CompoundFile, options);
            presentation.Parse(package, options);
            recordBudget.ThrowIfExceeded();
            decodedStorageBudget.ThrowIfExceeded();
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
                ToBoundedOffset(documentOffset, documentStream.Length,
                    "document persist object"), options, _recordBudget);
            if (document.Type != RecordDocument) {
                throw new InvalidDataException($"Persist object {documentPersistId} is record 0x{document.Type:X4}, not DocumentContainer.");
            }

            LegacyPptRecord? documentAtom = document.Children.FirstOrDefault(record => record.Type == RecordDocumentAtom);
            ParseDocumentSettings(documentAtom);
            ParseDocumentHeaderFooterSettings(document, options);

            ParseBlipStore(document, package, options);
            ParsePictureBullets(document, options);
            ParseFontCollection(document, options);
            ParseSoundCollection(document, options);
            ParseNamedShows(document, options);
            ParseHyperlinks(document, options);
            ParseOleObjects(document, package, options);
            ValidateExternalObjectIdSeed(document, options);
            ParseVbaProject(document, package, options);

            ParseSpecialMasters(documentAtom, documentStream, persistOffsets, options);
            ParseMasters(document, documentStream, persistOffsets, options);

            LegacyPptRecord? slideList = document.Children.FirstOrDefault(record =>
                record.Type == RecordSlideListWithText && record.Instance == 0);
            if (slideList == null) {
                AddDiagnostic("PPT-SLIDES-MISSING", LegacyPptDiagnosticSeverity.Warning,
                    "The document has no slide list.", document.Offset);
                ValidateCustomShowSlideReferences();
                ValidateSoundReferences();
                ValidateExternalObjectSlideReferences();
                return;
            }
            IReadOnlyDictionary<uint, LegacyPptNotesDirectoryEntry> notesDirectory =
                ReadNotesDirectory(document, options);

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
                    ToBoundedOffset(slideOffset, documentStream.Length,
                        "slide persist object"), options, _recordBudget);
                if (slideRecord.Type != RecordSlide) {
                    AddDiagnostic("PPT-SLIDE-TYPE", LegacyPptDiagnosticSeverity.Warning,
                        $"Slide {slideId} points to record 0x{slideRecord.Type:X4}; the slide was skipped.", slideRecord.Offset);
                    continue;
                }

                var slide = new LegacyPptSlide(slideId, persistId) { Name = $"Slide {++slideIndex}" };
                ParseSlide(slideRecord, slide, options);
                TryReadNotes(slide, documentStream, persistOffsets, notesDirectory, options);
                _slides.Add(slide);
            }
            ValidateCustomShowSlideReferences();
            ValidateSoundReferences();
            ValidateExternalObjectSlideReferences();
        }

        private void ParseSlide(LegacyPptRecord slideRecord, LegacyPptSlide slide, LegacyPptImportOptions options) {
            ParseSlideAtom(slideRecord, slide, options);
            slide.HeaderFooter = ReadHeaderFooterSettings(slideRecord, instance: 0,
                $"slide {slide.SlideId}", allowHeader: false, options);
            slide.RoundTripTheme = ReadRoundTripTheme(slideRecord,
                $"slide {slide.SlideId}", options);
            slide.ColorScheme = ReadColorScheme(slideRecord);
            ParseSlideShowInfo(slideRecord, slide, options);
            ParseComments(slideRecord, slide, options);
            LegacyPptColorScheme? effectiveScheme = slide.FollowsMasterColorScheme
                ? _masters.FirstOrDefault(master => master.MasterId == slide.MasterId)?.ColorScheme
                : slide.ColorScheme;
            slide.Background = ReadBackground(slideRecord,
                effectiveScheme ?? slide.ColorScheme, options);
            ParseShapes(slideRecord, slide.AddShape, "slide", options,
                effectiveScheme ?? slide.ColorScheme, slide.AddConnectorRule);
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
                    ToBoundedOffset(masterOffset, documentStream.Length,
                        "master persist object"), options, _recordBudget);
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
                ParseMasterSlideAtom(masterRecord, master, slideAtom, options);
                master.HeaderFooter = ReadHeaderFooterSettings(masterRecord, instance: 0,
                    $"master 0x{masterId:X8}", allowHeader: false, options);
                master.RoundTripTheme = ReadRoundTripTheme(masterRecord,
                    $"master 0x{masterId:X8}", options);
                master.ColorScheme = ReadColorScheme(masterRecord);
                LegacyPptColorScheme? effectiveScheme = !isMainMaster && master.FollowsMasterColorScheme
                    ? _masters.FirstOrDefault(candidate => candidate.MasterId == master.ParentMasterId)?.ColorScheme
                    : master.ColorScheme;
                master.Background = ReadBackground(masterRecord,
                    effectiveScheme ?? master.ColorScheme, options);
                ParseTextMasterStyles(masterRecord, master, effectiveScheme ?? master.ColorScheme, options);
                ParseShapes(masterRecord, master.AddShape, isMainMaster ? "main master" : "title master",
                    options, effectiveScheme ?? master.ColorScheme, master.AddConnectorRule);
                _masters.Add(master);
            }
        }

        private void ParseTextMasterStyles(LegacyPptRecord masterRecord, LegacyPptMaster master,
            LegacyPptColorScheme? colorScheme, LegacyPptImportOptions options) {
            IReadOnlyDictionary<ushort, LegacyPptRecord> style9Records =
                ReadMasterTextStyle9Records(masterRecord, options);
            var consumedStyle9Types = new HashSet<ushort>();
            foreach (LegacyPptRecord record in masterRecord.Children.Where(child =>
                         child.Type == RecordTextMasterStyleAtom)) {
                LegacyPptTextMasterStyle? style = LegacyPptTextMasterStyleReader.Read(record,
                    colorScheme, _fontsByIndex);
                if (style == null) {
                    if (options.ReportUnsupportedContent) {
                        AddDiagnostic("PPT-TEXT-MASTER-STYLE-TYPE",
                            LegacyPptDiagnosticSeverity.Warning,
                            $"TextMasterStyleAtom instance {record.Instance} is not a defined text type and remains preserve-only.",
                            record.Offset);
                    }
                    continue;
                }
                if (style9Records.TryGetValue(record.Instance,
                        out LegacyPptRecord? style9)) {
                    style = LegacyPptTextMasterStyleReader.ApplyStyle9(
                        style, style9, _pictureBulletsByIndex);
                    consumedStyle9Types.Add(record.Instance);
                }
                master.AddTextMasterStyle(style);
                if (style.IsTruncated) {
                    AddDiagnostic("PPT-TEXT-MASTER-STYLE-TRUNCATED",
                        LegacyPptDiagnosticSeverity.Warning,
                        $"The {style.TextType} TextMasterStyleAtom is malformed or truncated and remains preserve-only.",
                        record.Offset);
                } else if (style.HasUnprojectedFormatting && options.ReportUnsupportedContent) {
                    AddDiagnostic("PPT-TEXT-MASTER-STYLE-PARTIAL",
                        LegacyPptDiagnosticSeverity.Warning,
                        $"The {style.TextType} master text style contains legacy-only formatting that remains preserve-only.",
                        record.Offset);
                }
            }
            foreach (ushort instance in style9Records.Keys.Where(instance =>
                         !consumedStyle9Types.Contains(instance))) {
                AddDiagnostic("PPT-TEXT-MASTER-STYLE9-ORPHAN",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"TextMasterStyle9Atom instance {instance} has no base TextMasterStyleAtom and remains preserve-only.",
                    style9Records[instance].Offset);
            }
        }

        private void ParseShapes(LegacyPptRecord ownerRecord, Action<LegacyPptShape> addShape,
            string ownerDescription, LegacyPptImportOptions options, LegacyPptColorScheme? colorScheme,
            Action<LegacyPptConnectorRule>? addConnectorRule = null) {
            LegacyPptRecord? drawing = ownerRecord.Children.FirstOrDefault(record => record.Type == RecordDrawing);
            if (drawing == null) return;

            LegacyPptRecord? primaryShapeGroup = drawing.DescendantsAndSelf()
                .FirstOrDefault(record => record.Type == OfficeArtSpgrContainer);
            if (primaryShapeGroup == null) return;

            foreach (LegacyPptRecord child in primaryShapeGroup.Children) {
                LegacyPptShape? shape = child.Type switch {
                    OfficeArtSpContainer => ParseShapeContainer(child, ownerDescription, options, colorScheme),
                    OfficeArtSpgrContainer => ParseShapeGroup(child, ownerDescription, options, colorScheme),
                    _ => null
                };
                if (shape != null) addShape(shape);
            }
            if (addConnectorRule != null) {
                ParseConnectorRules(drawing, addConnectorRule, options);
            }
        }

        private void ParseConnectorRules(LegacyPptRecord drawing,
            Action<LegacyPptConnectorRule> addConnectorRule, LegacyPptImportOptions options) {
            foreach (LegacyPptRecord solver in drawing.DescendantsAndSelf().Where(record =>
                         record.Type == OfficeArtSolverContainer)) {
                foreach (LegacyPptRecord rule in solver.Children.Where(record =>
                             record.Type == OfficeArtFConnectorRule)) {
                    if (rule.PayloadLength < 24) {
                        if (options.ReportUnsupportedContent) {
                            AddDiagnostic("PPT-CONNECTOR-RULE-TRUNCATED",
                                LegacyPptDiagnosticSeverity.Warning,
                                "An OfficeArt connector rule is truncated and was skipped.", rule.Offset);
                        }
                        continue;
                    }
                    addConnectorRule(new LegacyPptConnectorRule(
                        rule.ReadUInt32(0), rule.ReadUInt32(4), rule.ReadUInt32(8),
                        rule.ReadUInt32(12), rule.ReadUInt32(16), rule.ReadUInt32(20)));
                }
            }
        }

        private LegacyPptShape? ParseShapeContainer(LegacyPptRecord shapeContainer,
            string ownerDescription, LegacyPptImportOptions options, LegacyPptColorScheme? colorScheme) {
            LegacyPptRecord? fsp = shapeContainer.Children.FirstOrDefault(record => record.Type == OfficeArtFsp);
            LegacyPptRecord? anchor = shapeContainer.Children.FirstOrDefault(record =>
                record.Type == OfficeArtClientAnchor || record.Type == OfficeArtChildAnchor);
            if (fsp == null || anchor == null || fsp.PayloadLength < 8) return null;

            ushort shapeType = fsp.Instance;
            uint shapeId = fsp.ReadUInt32(0);
            uint shapeFlags = fsp.ReadUInt32(4);
            if ((shapeFlags & OfficeArtBackgroundShapeFlag) != 0) return null;
            LegacyPptBounds bounds;
            try {
                bounds = ReadBounds(anchor);
            } catch (InvalidDataException) {
                AddDiagnostic("PPT-SHAPE-ANCHOR", LegacyPptDiagnosticSeverity.Warning,
                    "A shape has an unsupported or truncated anchor and was skipped.", anchor.Offset);
                return null;
            }

            LegacyPptRecord? textbox = shapeContainer.Children.FirstOrDefault(record =>
                record.Type == OfficeArtClientTextbox);
            LegacyPptTextData textData = textbox == null
                ? new LegacyPptTextData(string.Empty, 0)
                : ReadText(textbox);
            string text = textData.Text;
            LegacyPptRecord? textStyle = textbox?.DescendantsAndSelf().FirstOrDefault(record =>
                record.Type == RecordStyleTextPropAtom);
            LegacyPptRecord? textHeader = textbox?.DescendantsAndSelf().FirstOrDefault(record =>
                record.Type == RecordTextHeaderAtom);
            LegacyPptTextType? textType = ReadTextType(textHeader);
            if (textHeader != null && !textType.HasValue && options.ReportUnsupportedContent) {
                AddDiagnostic("PPT-TEXT-TYPE-INVALID", LegacyPptDiagnosticSeverity.Warning,
                    "A TextHeaderAtom is truncated or contains an undefined text type; its text remains available without master-style classification.",
                    textHeader.Offset);
            }
            LegacyPptRecord? textRulerRecord = textbox?.DescendantsAndSelf().FirstOrDefault(record =>
                record.Type == RecordTextRulerAtom);
            LegacyPptTextRuler? textRuler = LegacyPptTextRulerReader.Read(textRulerRecord,
                out bool isTextRulerTruncated);
            LegacyPptTextBody textBody = LegacyPptTextStyleReader.Read(text, textData.RawCharacterCount,
                textStyle, colorScheme, _fontsByIndex, textType, textRuler,
                hasRulerRecord: textRulerRecord != null, isRulerTruncated: isTextRulerTruncated);
            LegacyPptRecord? textStyle9 = ReadShapeStyle9(shapeContainer,
                options, out bool isTextStyle9Malformed);
            textBody = LegacyPptTextStyle9Reader.Apply(textBody, textStyle9,
                    _pictureBulletsByIndex, isTextStyle9Malformed);
            LegacyPptRecord? textSpecialInfo = textbox?.DescendantsAndSelf()
                .FirstOrDefault(record =>
                    record.Type == RecordTextSpecialInfoAtom);
            bool hasDuplicateTextSpecialInfo = textbox?.DescendantsAndSelf()
                .Count(record => record.Type == RecordTextSpecialInfoAtom) > 1;
            textBody = hasDuplicateTextSpecialInfo
                ? textBody.WithLanguageInformation(
                    Array.Empty<LegacyPptTextLanguageRun>(),
                    hasTextSpecialInfoRecord: true,
                    hasUnprojectedTextSpecialInfo: true,
                    isTextSpecialInfoTruncated: true)
                : LegacyPptTextSpecialInfoCodec.Apply(textBody,
                    textSpecialInfo, textData.RawCharacterCount);
            IReadOnlyList<LegacyPptTextField> fields = ReadTextFields(
                textbox, text, options, out bool hasFieldRecords,
                out bool isFieldDataMalformed);
            textBody = textBody.WithFields(fields, hasFieldRecords,
                    isFieldDataMalformed)
                .WithInteractions(ReadTextInteractions(textbox, text.Length, options));
            if (textBody.IsTextSpecialInfoTruncated
                && options.ReportUnsupportedContent) {
                AddDiagnostic("PPT-TEXT-SPECIAL-INFO-TRUNCATED",
                    LegacyPptDiagnosticSeverity.Warning,
                    "A TextSpecialInfoAtom is malformed or does not cover its text exactly; language metadata remains preserve-only.",
                    textSpecialInfo?.Offset ?? shapeContainer.Offset);
            } else if (textBody.HasUnprojectedTextSpecialInfo
                       && options.ReportUnsupportedContent) {
                AddDiagnostic("PPT-TEXT-SPECIAL-INFO-PARTIAL",
                    LegacyPptDiagnosticSeverity.Information,
                    "Text language and spelling metadata was projected while grammar, bidirectional analysis, extension, smart-tag, transient language, or mixed explicit no-language fields remain preserve-only.",
                    textSpecialInfo?.Offset ?? shapeContainer.Offset);
            }
            LegacyPptPlaceholder? placeholder = ReadPlaceholder(shapeContainer, options);
            LegacyPptShapeKind kind = ClassifyShape(shapeType, textbox != null || text.Length > 0,
                (shapeFlags & (1U << 8)) != 0);
            LegacyPptRecord? fopt = shapeContainer.Children.FirstOrDefault(record =>
                record.Type == OfficeArtFopt);
            LegacyPptRecord? tertiaryFopt = shapeContainer.Children
                .FirstOrDefault(record =>
                    record.Type == OfficeArtTertiaryFopt);
            OfficeArtShapeStyle style = ReadShapeStyle(fopt, tertiaryFopt);
            OfficeArtPictureProperties pictureProperties =
                OfficeArtPictureProperties.Decode(style.Properties);
            OfficeArtShapeTransform transform = OfficeArtShapeTransform.Decode(shapeFlags,
                style.Properties);
            int? pictureStoreIndex = ReadPictureStoreIndex(style);
            OfficeArtBlipStoreEntry? picture = ResolvePicture(pictureStoreIndex);
            ReadShapeExternalObject(shapeContainer, options,
                out LegacyPptEmbeddedOleObject? oleObject,
                out LegacyPptLinkedOleObject? linkedOleObject,
                out LegacyPptActiveXControl? activeXControl,
                out LegacyPptMedia? media);
            bool isPictureFrame = shapeType == 75;
            if (oleObject != null) {
                kind = LegacyPptShapeKind.OleObject;
            } else if (media?.HasProjectableAudio == true) {
                kind = LegacyPptShapeKind.Media;
            } else if (isPictureFrame) {
                kind = picture?.HasImportableImage == true
                    ? LegacyPptShapeKind.Picture
                    : LegacyPptShapeKind.Unsupported;
            }
            var shape = new LegacyPptShape(kind, shapeType, shapeId, shapeContainer.Offset,
                bounds, text, placeholder, style,
                ResolveShapeColor(style.FillColor, colorScheme),
                ResolveShapeColor(style.LineColor, colorScheme), pictureStoreIndex, picture,
                transform, shadowColor: ResolveShapeColor(style.ShadowColor, colorScheme),
                textBody: textBody, interactions: ReadShapeInteractions(shapeContainer, options),
                animation: ReadShapeAnimation(shapeContainer, options),
                oleObject: oleObject, linkedOleObject: linkedOleObject,
                activeXControl: activeXControl, media: media,
                pictureTransparentColor: ResolveShapeColor(
                    pictureProperties.TransparentColor, colorScheme),
                pictureRecolorColor: ResolveShapeColor(
                    pictureProperties.RecolorColor, colorScheme),
                fillBackColor: ResolveShapeColor(style.FillBackColor, colorScheme),
                fillGradientStops: ResolveShapeGradientStops(style, colorScheme));

            if (textBody.IsStyleTruncated) {
                AddDiagnostic("PPT-TEXT-STYLE-TRUNCATED", LegacyPptDiagnosticSeverity.Warning,
                    "A StyleTextPropAtom is malformed or truncated; its text remains available as plain text.",
                    textStyle?.Offset ?? textbox?.Offset);
            } else {
                if (textBody.IsStyle9Truncated
                    && !isTextStyle9Malformed) {
                    AddDiagnostic("PPT-TEXT-STYLE9-TRUNCATED",
                        LegacyPptDiagnosticSeverity.Warning,
                        "A StyleTextProp9Atom is malformed or cannot be linked to its base character runs; its extended formatting remains preserve-only.",
                        textStyle9?.Offset ?? shapeContainer.Offset);
                }
                if (textBody.HasUnprojectedParagraphFormatting) {
                    AddDiagnostic("PPT-TEXT-PARAGRAPH-PARTIAL", LegacyPptDiagnosticSeverity.Warning,
                        "A text ruler field, unresolved bullet resource, or unsupported bullet size is preserved but not projected yet.",
                        textStyle?.Offset ?? textbox?.Offset);
                }
                if (textBody.HasUnprojectedCharacterFormatting) {
                    AddDiagnostic("PPT-TEXT-CHARACTER-PARTIAL", LegacyPptDiagnosticSeverity.Warning,
                        "Typeface references and legacy-only character effects are preserved but not projected yet.",
                        textStyle?.Offset ?? textbox?.Offset);
                }
            }
            if (textBody.IsRulerTruncated) {
                AddDiagnostic("PPT-TEXT-RULER-TRUNCATED", LegacyPptDiagnosticSeverity.Warning,
                    "A TextRulerAtom is malformed or truncated; its ruler formatting remains preserved only.",
                    textRulerRecord?.Offset ?? textbox?.Offset);
            }
            if (fopt != null && fopt.Instance > 0 && style.Properties.Count == 0) {
                AddDiagnostic("PPT-SHAPE-STYLE-TRUNCATED", LegacyPptDiagnosticSeverity.Warning,
                    "The OfficeArt property table is truncated and could not be decoded.", fopt.Offset);
            } else if (!isPictureFrame && style.HasUnprojectedVisualStyle) {
                AddDiagnostic("PPT-SHAPE-STYLE-PARTIAL", LegacyPptDiagnosticSeverity.Warning,
                    "The shape uses visual styling that is only partially projected; exact source properties remain preserved.",
                    shapeContainer.Offset);
            }

            if (isPictureFrame && oleObject == null
                && linkedOleObject == null && activeXControl == null
                && media == null
                && kind == LegacyPptShapeKind.Unsupported
                && options.ReportUnsupportedContent) {
                AddPictureDiagnostic(pictureStoreIndex, picture, shapeContainer.Offset);
            } else if (linkedOleObject != null
                       && options.ReportUnsupportedContent) {
                AddDiagnostic("PPT-OLE-LINK-PRESERVED",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"Linked OLE identifier {linkedOleObject.Id} and its cache are preserved exactly but are not projected to an editable Open XML object.",
                    shapeContainer.Offset);
            } else if (activeXControl != null
                       && options.ReportUnsupportedContent) {
                AddDiagnostic("PPT-ACTIVEX-PRESERVED",
                    LegacyPptDiagnosticSeverity.Warning,
                    $"ActiveX identifier {activeXControl.Id} and its Office Forms storage are preserved exactly but are not projected to an editable Open XML control.",
                    shapeContainer.Offset);
            } else if (media != null && options.ReportUnsupportedContent
                       && (!media.HasProjectableAudio || media.Loop
                           || media.Rewind || media.Narration)) {
                AddDiagnostic("PPT-MEDIA-PRESERVED",
                    LegacyPptDiagnosticSeverity.Warning,
                    media.HasProjectableAudio
                        ? $"Embedded WAV media {media.Id} is editable, while its loop, rewind, or narration playback flag remains preserve-only."
                        : $"{media.Kind} media {media.Id} is retained with its native metadata but is not projected as editable embedded media.",
                    shapeContainer.Offset);
            } else if (kind == LegacyPptShapeKind.Unsupported && options.ReportUnsupportedContent) {
                AddDiagnostic("PPT-SHAPE-UNSUPPORTED", LegacyPptDiagnosticSeverity.Warning,
                    $"OfficeArt shape type {shapeType} on a {ownerDescription} is not projected to the editable PowerPoint model.", shapeContainer.Offset);
            } else if (LegacyPptShapeGeometryMapper.IsApproximation(shapeType)
                && options.ReportUnsupportedContent) {
                AddDiagnostic("PPT-SHAPE-GEOMETRY-APPROXIMATED", LegacyPptDiagnosticSeverity.Warning,
                    $"OfficeArt shape type {shapeType} uses the closest DrawingML preset geometry.",
                    shapeContainer.Offset);
            }
            return shape;
        }

        private LegacyPptShape? ParseShapeGroup(LegacyPptRecord groupContainer,
            string ownerDescription, LegacyPptImportOptions options, LegacyPptColorScheme? colorScheme) {
            LegacyPptRecord? descriptor = groupContainer.Children.FirstOrDefault(record =>
                record.Type == OfficeArtSpContainer
                && record.Children.Any(child => child.Type == OfficeArtFspgr));
            LegacyPptRecord? fspgr = descriptor?.Children.FirstOrDefault(record => record.Type == OfficeArtFspgr);
            LegacyPptRecord? fsp = descriptor?.Children.FirstOrDefault(record => record.Type == OfficeArtFsp);
            LegacyPptRecord? fopt = descriptor?.Children.FirstOrDefault(record => record.Type == OfficeArtFopt);
            LegacyPptRecord? tertiaryFopt = descriptor?.Children
                .FirstOrDefault(record =>
                    record.Type == OfficeArtTertiaryFopt);
            LegacyPptRecord? anchor = descriptor?.Children.FirstOrDefault(record =>
                record.Type == OfficeArtClientAnchor || record.Type == OfficeArtChildAnchor);
            if (descriptor == null || fspgr == null || fsp == null || anchor == null
                || fsp.PayloadLength < 8 || fspgr.PayloadLength < 16) {
                if (options.ReportUnsupportedContent) {
                    AddDiagnostic("PPT-GROUP-TRUNCATED", LegacyPptDiagnosticSeverity.Warning,
                        $"A shape group on a {ownerDescription} has incomplete coordinate or anchor records.",
                        groupContainer.Offset);
                }
                return null;
            }

            LegacyPptBounds bounds;
            LegacyPptBounds coordinateBounds;
            try {
                bounds = ReadBounds(anchor);
                coordinateBounds = CreateBounds(fspgr.ReadInt32(0), fspgr.ReadInt32(4),
                    fspgr.ReadInt32(8), fspgr.ReadInt32(12));
            } catch (InvalidDataException) {
                AddDiagnostic("PPT-GROUP-ANCHOR", LegacyPptDiagnosticSeverity.Warning,
                    "A shape group has an unsupported or truncated coordinate system.",
                    groupContainer.Offset);
                return null;
            }

            var children = new List<LegacyPptShape>();
            foreach (LegacyPptRecord child in groupContainer.Children.Where(record =>
                         !ReferenceEquals(record, descriptor))) {
                LegacyPptShape? shape = child.Type switch {
                    OfficeArtSpContainer => ParseShapeContainer(child, ownerDescription, options, colorScheme),
                    OfficeArtSpgrContainer => ParseShapeGroup(child, ownerDescription, options, colorScheme),
                    _ => null
                };
                if (shape != null) children.Add(shape);
            }
            if (children.Count == 0) return null;
            OfficeArtShapeStyle style = ReadShapeStyle(fopt, tertiaryFopt);
            OfficeArtShapeTransform transform = OfficeArtShapeTransform.Decode(fsp.ReadUInt32(4),
                style.Properties);
            return new LegacyPptShape(LegacyPptShapeKind.Group, fsp.Instance, fsp.ReadUInt32(0),
                groupContainer.Offset, bounds, string.Empty, placeholder: null,
                style, ResolveShapeColor(style.FillColor, colorScheme),
                ResolveShapeColor(style.LineColor, colorScheme), transform: transform,
                groupCoordinateBounds: coordinateBounds, children: children,
                tableStyleFlags: ReadOfficeImoTableStyleFlags(descriptor,
                    options),
                shadowColor: ResolveShapeColor(style.ShadowColor, colorScheme),
                interactions: ReadShapeInteractions(descriptor, options),
                animation: ReadShapeAnimation(descriptor, options),
                fillBackColor: ResolveShapeColor(style.FillBackColor, colorScheme),
                fillGradientStops: ResolveShapeGradientStops(style, colorScheme));
        }

        private static IReadOnlyList<LegacyPptGradientStop> ResolveShapeGradientStops(
            OfficeArtShapeStyle style, LegacyPptColorScheme? colorScheme) =>
            style.FillGradientStops.Select(stop => new LegacyPptGradientStop(
                ResolveShapeColor(stop.Color, colorScheme), stop.Position)).ToArray();

        private void ParseBlipStore(LegacyPptRecord document, LegacyPptPackage package,
            LegacyPptImportOptions options) {
            LegacyPptRecord? store = document.DescendantsAndSelf()
                .FirstOrDefault(record => record.Type == OfficeArtBStoreContainer);
            if (store == null) return;
            foreach (LegacyPptRecord fbse in store.Children.Where(record => record.Type == OfficeArtFbse)) {
                byte[] bytes = fbse.CopyRecordBytes();
                if (OfficeArtBlipStoreEntryReader.TryRead(bytes, 8, fbse.PayloadLength, fbse.Instance,
                        package.PicturesStream, out OfficeArtBlipStoreEntry? entry,
                        options.MaxInputBytes) && entry != null) {
                    _blipStoreEntries.Add(entry);
                } else if (options.ReportUnsupportedContent) {
                    AddDiagnostic("PPT-PICTURE-FBSE-TRUNCATED", LegacyPptDiagnosticSeverity.Warning,
                        "An OfficeArt picture-store entry is truncated and could not be decoded.", fbse.Offset);
                }
            }
        }

        private static int? ReadPictureStoreIndex(OfficeArtShapeStyle style) {
            OfficeArtProperty? property = style.Properties.FirstOrDefault(candidate =>
                candidate.PropertyId == 0x0104 && candidate.IsBlipId && candidate.Value > 0);
            if (property == null || property.Value > int.MaxValue) return null;
            return unchecked((int)property.Value);
        }

        private OfficeArtBlipStoreEntry? ResolvePicture(int? oneBasedIndex) {
            if (!oneBasedIndex.HasValue) return null;
            int index = oneBasedIndex.Value - 1;
            return (uint)index < (uint)_blipStoreEntries.Count ? _blipStoreEntries[index] : null;
        }

        private void AddPictureDiagnostic(int? storeIndex, OfficeArtBlipStoreEntry? picture,
            long offset) {
            if (picture == null) {
                AddDiagnostic("PPT-PICTURE-BLIP-MISSING", LegacyPptDiagnosticSeverity.Warning,
                    storeIndex.HasValue
                        ? $"The picture frame references missing BLIP store entry {storeIndex.Value}."
                        : "The picture frame has no valid BLIP store reference.", offset);
            } else if (picture.IsPayloadTruncated) {
                AddDiagnostic("PPT-PICTURE-BLIP-TRUNCATED", LegacyPptDiagnosticSeverity.Warning,
                    $"The {picture.BlipRecordTypeName ?? picture.RecordInstanceBlipTypeName} payload is truncated.",
                    offset);
            } else {
                AddDiagnostic("PPT-PICTURE-FORMAT-UNSUPPORTED", LegacyPptDiagnosticSeverity.Warning,
                    $"The {picture.BlipRecordTypeName ?? picture.RecordInstanceBlipTypeName} picture cannot be projected to an editable Open XML image part.",
                    offset);
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
                bool isChildAnchor = anchor.Type == OfficeArtChildAnchor;
                int left = anchor.ReadInt32(isChildAnchor ? 0 : 4);
                int top = anchor.ReadInt32(isChildAnchor ? 4 : 0);
                int right = anchor.ReadInt32(8);
                int bottom = anchor.ReadInt32(12);
                return CreateBounds(left, top, right, bottom);
            }
            throw new InvalidDataException("The OfficeArt anchor is too short.");
        }

        private static OfficeArtShapeStyle ReadShapeStyle(
            params LegacyPptRecord?[] propertyTables) {
            var properties = new List<OfficeArtProperty>();
            foreach (LegacyPptRecord? table in propertyTables) {
                if (table == null || table.PayloadLength == 0
                    || table.Instance == 0) continue;
                byte[] recordBytes = table.CopyRecordBytes();
                properties.AddRange(OfficeArtPropertyTableReader.Read(
                    recordBytes, 8, table.PayloadLength, table.Instance));
            }
            return OfficeArtShapeStyle.Decode(properties);
        }

        private static string? ResolveShapeColor(OfficeArtColorReference? reference,
            LegacyPptColorScheme? colorScheme) {
            if (!reference.HasValue) return null;
            Func<byte, OfficeColor?>? resolver = colorScheme == null
                ? null
                : colorScheme.ResolveOfficeArtColor;
            return reference.Value.TryResolve(resolver, out OfficeColor color)
                ? color.ToRgbHex()
                : null;
        }

        private static LegacyPptColorScheme? ReadColorScheme(LegacyPptRecord ownerRecord) {
            LegacyPptRecord? atom = ownerRecord.Children.LastOrDefault(record =>
                record.Type == RecordColorSchemeAtom && record.Instance == 1 && record.PayloadLength >= 32);
            if (atom == null) return null;
            var colors = new string[8];
            for (int index = 0; index < colors.Length; index++) {
                int offset = index * 4;
                colors[index] = string.Concat(
                    atom.ReadByte(offset).ToString("X2"),
                    atom.ReadByte(offset + 1).ToString("X2"),
                    atom.ReadByte(offset + 2).ToString("X2"));
            }
            return new LegacyPptColorScheme(colors);
        }

        private static LegacyPptBounds CreateBounds(int left, int top, int right, int bottom) =>
            new LegacyPptBounds(left, top, Math.Max(0, right - left), Math.Max(0, bottom - top));

        private static LegacyPptTextData ReadText(LegacyPptRecord textbox) {
            LegacyPptRecord? textRecord = textbox.DescendantsAndSelf().FirstOrDefault(record =>
                record.Type == RecordTextChars || record.Type == RecordTextBytes);
            if (textRecord == null) return new LegacyPptTextData(string.Empty, 0);
            string text = textRecord.Type == RecordTextChars
                ? textRecord.ReadUtf16Text()
                : textRecord.ReadLowByteUnicodeText();
            string rawText = text.TrimEnd('\0');
            return new LegacyPptTextData(rawText.Replace("\r", "\n"),
                rawText.Length);
        }

        private static LegacyPptTextType? ReadTextType(LegacyPptRecord? header) {
            if (header == null || header.PayloadLength < 4) return null;
            uint value = header.ReadUInt32(0);
            return value == 0 || value == 1 || value == 2 || value == 4
                || value == 5 || value == 6 || value == 7 || value == 8
                ? (LegacyPptTextType)value
                : null;
        }

        private readonly struct LegacyPptTextData {
            internal LegacyPptTextData(string text, int rawCharacterCount) {
                Text = text;
                RawCharacterCount = rawCharacterCount;
            }

            internal string Text { get; }

            internal int RawCharacterCount { get; }
        }

        private static LegacyPptShapeKind ClassifyShape(ushort shapeType, bool hasText,
            bool isConnector = false) {
            if (hasText || shapeType == 202) return LegacyPptShapeKind.TextBox;
            if (shapeType == 75) return LegacyPptShapeKind.Picture;
            if (isConnector) return LegacyPptShapeKind.Connector;
            if (shapeType == 1) return LegacyPptShapeKind.Rectangle;
            if (shapeType == 3) return LegacyPptShapeKind.Ellipse;
            if (shapeType == 20) return LegacyPptShapeKind.Line;
            if (LegacyPptShapeGeometryMapper.IsConnector(shapeType)) return LegacyPptShapeKind.Connector;
            if (LegacyPptShapeGeometryMapper.TryGetPreset(shapeType, out _)) return LegacyPptShapeKind.AutoShape;
            return LegacyPptShapeKind.Unsupported;
        }

        private void AddCompoundFeatureDiagnostics(OfficeCompoundFile compound, LegacyPptImportOptions options) {
            if (!options.ReportUnsupportedContent) return;
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
