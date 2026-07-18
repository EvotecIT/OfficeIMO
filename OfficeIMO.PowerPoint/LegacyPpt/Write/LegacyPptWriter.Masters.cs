using System.Collections.ObjectModel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        internal const int MaxNativeMasterCount = 4093;
        private const uint FirstMasterId = 0x80000000U;
        private const uint FirstMasterPersistId = 2U;
        private const ushort RecordHandoutForWrite = 0x0FC9;

        private static LegacyPptWriterMasterCatalog ReadMasterCatalog(
            PowerPointPresentation presentation,
            LegacyPptRecord templateDocument,
            IReadOnlyList<LegacyPptRecord> prototypes,
            LegacyPptRecord notesMasterPrototype,
            LegacyPptWriterTopology topology,
            LegacyPptWriterPictureBulletCatalog pictureBullets,
            LegacyPptWriterPictureCatalog pictureCatalog) {
            SlideMasterPart[] masterParts = presentation.OpenXmlDocument.PresentationPart?
                .SlideMasterParts.ToArray() ?? Array.Empty<SlideMasterPart>();
            if (masterParts.Length == 0) {
                throw new InvalidDataException(
                    "The Open XML presentation has no slide master to encode.");
            }
            if (masterParts.Length > MaxNativeMasterCount) {
                throw new NotSupportedException(
                    $"The binary PowerPoint persist-directory format supports at most {MaxNativeMasterCount} slide masters; "
                    + $"the presentation contains {masterParts.Length}.");
            }
            if (prototypes.Count == 0) {
                throw new InvalidDataException(
                    "The embedded binary PowerPoint template has no slide-master prototype.");
            }
            if (!TryReadMasterTextStyles(presentation, templateDocument,
                    pictureBullets,
                    out LegacyPptWriterMasterTextStyleCatalog textStyles,
                    out string? textStyleReason)) {
                throw new NotSupportedException(textStyleReason);
            }

            var masterIds = new Dictionary<string, uint>(StringComparer.Ordinal);
            var persistObjects = new List<byte[]>(topology.MasterSlotCount);
            var drawingShapeCounts = new Dictionary<uint, int>();
            for (int index = 0; index < topology.MasterSlotCount; index++) {
                LegacyPptRecord prototype = prototypes[index % prototypes.Count];
                if (index < masterParts.Length) {
                    SlideMasterPart source = masterParts[index];
                    if (!TryReadBackground(source,
                            out LegacyPptWriterBackground? background,
                            out string? backgroundReason)) {
                        throw new NotSupportedException(backgroundReason);
                    }
                    IReadOnlyList<PowerPointShape> sourceShapes =
                        ReadMasterShapesForWrite(source, out _);
                    IReadOnlyList<PowerPointShape> supportedShapes = sourceShapes
                        .Where(IsSupportedMasterShape).ToArray();
                    uint drawingId = topology.GetMasterDrawingId(index);
                    IReadOnlyList<byte[]> roundTripThemeRecords =
                        BuildRoundTripThemeRecords(source.ThemePart?.Theme,
                            source.SlideMaster?.ColorMap);
                    persistObjects.Add(BuildMasterRecord(prototype,
                        ReadColorScheme(source.ThemePart), background,
                        supportedShapes, drawingId,
                        LegacyPptWriterShapeContext.MainMaster,
                        textStyles.Get(source),
                        textStyles.GetStyle9(source), roundTripThemeRecords,
                        fonts: textStyles.Fonts,
                        pictureBullets: pictureBullets,
                        pictureCatalog: pictureCatalog));
                    drawingShapeCounts.Add(drawingId,
                        CountDrawingShapes(supportedShapes, pictureCatalog));
                    masterIds.Add(masterParts[index].Uri.ToString(),
                        checked(FirstMasterId + unchecked((uint)index)));
                } else {
                    persistObjects.Add(prototype.CopyRecordBytes());
                }
            }
            NotesMasterPart? notesMasterPart = presentation.OpenXmlDocument
                .PresentationPart?.NotesMasterPart;
            LegacyPptWriterBackground? notesBackground = null;
            if (notesMasterPart != null
                && !TryReadBackground(notesMasterPart, out notesBackground,
                    out string? notesBackgroundReason)) {
                throw new NotSupportedException(notesBackgroundReason);
            }
            IReadOnlyList<PowerPointShape>? notesShapes = notesMasterPart == null
                ? null
                : ReadMasterShapesForWrite(notesMasterPart, out _)
                    .Where(IsSupportedMasterShape).ToArray();
            uint notesDrawingId = topology.NotesMasterDrawingId;
            byte[] notesMaster = BuildMasterRecord(notesMasterPrototype,
                ReadColorScheme(notesMasterPart?.ThemePart
                    ?? masterParts[0].ThemePart), notesBackground,
                notesShapes, notesDrawingId,
                LegacyPptWriterShapeContext.NotesMaster,
                roundTripThemeRecords: BuildRoundTripThemeRecords(
                    notesMasterPart?.ThemePart?.Theme
                        ?? masterParts[0].ThemePart?.Theme,
                    notesMasterPart?.NotesMaster?.ColorMap),
                fonts: textStyles.Fonts,
                pictureBullets: pictureBullets,
                pictureCatalog: pictureCatalog);
            if (notesShapes != null) {
                drawingShapeCounts.Add(notesDrawingId,
                    CountDrawingShapes(notesShapes, pictureCatalog));
            }
            HandoutMasterPart? handoutMasterPart = presentation.OpenXmlDocument
                .PresentationPart?.HandoutMasterPart;
            byte[]? handoutMaster = null;
            if (handoutMasterPart != null) {
                if (!TryReadBackground(handoutMasterPart,
                        out LegacyPptWriterBackground? handoutBackground,
                        out string? handoutBackgroundReason)) {
                    throw new NotSupportedException(handoutBackgroundReason);
                }
                IReadOnlyList<PowerPointShape> handoutShapes =
                    ReadMasterShapesForWrite(handoutMasterPart, out _)
                        .Where(IsSupportedMasterShape).ToArray();
                handoutMaster = BuildHandoutMasterRecord(
                    notesMasterPrototype,
                    ReadColorScheme(handoutMasterPart.ThemePart
                        ?? masterParts[0].ThemePart), handoutBackground,
                    handoutShapes, topology.HandoutMasterDrawingId,
                    BuildRoundTripThemeRecords(
                        handoutMasterPart.ThemePart?.Theme
                            ?? masterParts[0].ThemePart?.Theme,
                        handoutMasterPart.HandoutMaster?.ColorMap),
                    textStyles.Fonts, pictureBullets, pictureCatalog);
                drawingShapeCounts.Add(topology.HandoutMasterDrawingId,
                    CountDrawingShapes(handoutShapes, pictureCatalog));
            }
            return new LegacyPptWriterMasterCatalog(masterIds, persistObjects,
                notesMaster, handoutMaster, masterParts.Length,
                drawingShapeCounts, textStyles.Fonts);
        }

        private static byte[] BuildHandoutMasterRecord(
            LegacyPptRecord drawingPrototypeOwner,
            LegacyPptWriterColorScheme scheme,
            LegacyPptWriterBackground? background,
            IReadOnlyList<PowerPointShape> shapes, uint drawingId,
            IReadOnlyList<byte[]> roundTripThemeRecords,
            LegacyPptWriterFontCatalog fonts,
            LegacyPptWriterPictureBulletCatalog pictureBullets,
            LegacyPptWriterPictureCatalog pictureCatalog) {
            LegacyPptRecord drawingPrototype = drawingPrototypeOwner.Children
                .First(record => record.Type == RecordDrawing);
            var interactions = new LegacyPptWriterInteractionCatalog();
            var animations = new LegacyPptWriterAnimationCatalog(
                new Dictionary<string, LegacyPptWriterAnimation>(
                    StringComparer.Ordinal));
            return BuildContainer(RecordHandoutForWrite, instance: 0,
                new[] {
                    BuildDrawingRecord(drawingPrototype, shapes, drawingId,
                        interactions, animations, fonts, background,
                        LegacyPptWriterShapeContext.HandoutMaster,
                        pictureCatalog: pictureCatalog,
                        pictureBullets: pictureBullets),
                    BuildColorSchemeAtom(scheme)
                }.Concat(roundTripThemeRecords));
        }

        private static byte[] BuildMasterRecord(LegacyPptRecord prototype,
            LegacyPptWriterColorScheme scheme,
            LegacyPptWriterBackground? background,
            IReadOnlyList<PowerPointShape>? shapes = null,
            uint? drawingId = null,
            LegacyPptWriterShapeContext shapeContext =
                LegacyPptWriterShapeContext.MainMaster,
            IReadOnlyList<byte[]>? textStyleRecords = null,
            IReadOnlyList<byte[]>? textStyle9Records = null,
            IReadOnlyList<byte[]>? roundTripThemeRecords = null,
            bool rewriteColorScheme = true,
            IReadOnlyList<int>? colorSchemeSlotsToRewrite = null,
            LegacyPptWriterFontCatalog? fonts = null,
            LegacyPptWriterPictureBulletCatalog? pictureBullets = null,
            LegacyPptWriterPictureCatalog? pictureCatalog = null) {
            var children = new List<byte[]>(prototype.Children.Count);
            bool wroteScheme = false;
            bool wroteTextStyles = false;
            var interactions = new LegacyPptWriterInteractionCatalog();
            var animations = new LegacyPptWriterAnimationCatalog(
                new Dictionary<string, LegacyPptWriterAnimation>(
                    StringComparer.Ordinal));
            foreach (LegacyPptRecord child in prototype.Children) {
                if (shapeContext == LegacyPptWriterShapeContext.NotesMaster
                    && child.Type == RecordNotesAtom
                    && child.PayloadLength >= 6) {
                    byte[] atom = child.CopyRecordBytes();
                    WriteUInt32(atom, 8, 0);
                    WriteUInt16(atom, 12, 0);
                    children.Add(atom);
                } else if (rewriteColorScheme && child.Type == RecordColorSchemeAtom
                    && child.Instance == 1) {
                    children.Add(colorSchemeSlotsToRewrite == null
                        ? BuildColorSchemeAtom(scheme)
                        : PatchColorSchemeAtom(child, scheme,
                            colorSchemeSlotsToRewrite));
                    wroteScheme = true;
                } else if (child.Type == RecordDrawing && drawingId.HasValue) {
                    if (shapes != null) {
                        children.Add(BuildDrawingRecord(child, shapes,
                            drawingId.Value, interactions, animations,
                            fonts ?? throw new InvalidOperationException(
                                "Master shape text requires the document font catalog."),
                            background,
                            shapeContext,
                            pictureCatalog: pictureCatalog,
                            pictureBullets: pictureBullets));
                    } else {
                        children.Add(RewriteDrawingId(child, drawingId.Value));
                    }
                } else if (background != null && child.Type == RecordDrawing) {
                    children.Add(BuildBackgroundDrawingRecord(child,
                        background, pictureCatalog));
                } else if (child.Type == RecordTextMasterStyleAtomForWrite
                           && textStyleRecords != null) {
                    if (!wroteTextStyles) {
                        children.AddRange(textStyleRecords);
                        wroteTextStyles = true;
                    }
                } else if (roundTripThemeRecords == null
                           || !IsRoundTripThemeRecord(child.Type)) {
                    children.Add(child.CopyRecordBytes());
                }
            }
            if (rewriteColorScheme && !wroteScheme) {
                children.Add(BuildColorSchemeAtom(scheme));
            }
            if (textStyleRecords != null && !wroteTextStyles) {
                children.AddRange(textStyleRecords);
            }
            if (roundTripThemeRecords != null) {
                children.AddRange(roundTripThemeRecords);
            }
            byte[] bytes = BuildContainer(prototype.Type, prototype.Instance,
                children);
            if (textStyle9Records == null) return bytes;
            LegacyPptRecord rewritten = LegacyPptRecordReader.ReadSingle(bytes,
                0, new LegacyPptImportOptions());
            if (!TryRewriteMasterTextStyle9Records(rewritten,
                    textStyle9Records, instancesToReplace: null,
                    replaceAllExisting: true, out bytes)) {
                throw new InvalidDataException(
                    "The binary master prototype has malformed or duplicate PPT9 programmable tags.");
            }
            return bytes;
        }

        /// <summary>
        /// Rewrites only the classic color scheme and DrawingML round-trip theme records of an imported master.
        /// Every unrelated source child record is retained byte-for-byte.
        /// </summary>
        internal static byte[] BuildPreservedMasterThemeRecord(
            LegacyPptRecord prototype, SlideMasterPart source,
            IReadOnlyList<int> changedClassicColorSlots) {
            if (prototype == null) throw new ArgumentNullException(nameof(prototype));
            if (source == null) throw new ArgumentNullException(nameof(source));
            return BuildPreservedMasterThemeRecord(prototype, source.ThemePart,
                source.SlideMaster?.ColorMap, changedClassicColorSlots);
        }

        internal static byte[] BuildPreservedMasterThemeRecord(
            LegacyPptRecord prototype, NotesMasterPart source,
            IReadOnlyList<int> changedClassicColorSlots) {
            if (prototype == null) throw new ArgumentNullException(nameof(prototype));
            if (source == null) throw new ArgumentNullException(nameof(source));
            return BuildPreservedMasterThemeRecord(prototype, source.ThemePart,
                source.NotesMaster?.ColorMap, changedClassicColorSlots);
        }

        internal static byte[] BuildPreservedMasterThemeRecord(
            LegacyPptRecord prototype, HandoutMasterPart source,
            IReadOnlyList<int> changedClassicColorSlots) {
            if (prototype == null) throw new ArgumentNullException(nameof(prototype));
            if (source == null) throw new ArgumentNullException(nameof(source));
            return BuildPreservedMasterThemeRecord(prototype, source.ThemePart,
                source.HandoutMaster?.ColorMap, changedClassicColorSlots);
        }

        internal static byte[] BuildPreservedMasterThemeRecord(
            LegacyPptRecord prototype, SlideLayoutPart source,
            IReadOnlyList<int> changedClassicColorSlots) {
            if (prototype == null) throw new ArgumentNullException(nameof(prototype));
            if (source == null) throw new ArgumentNullException(nameof(source));
            A.ThemeOverride theme = source.ThemeOverridePart?.ThemeOverride
                ?? throw new InvalidDataException(
                    "A projected binary title master has no DrawingML theme override to preserve.");
            if (changedClassicColorSlots == null) {
                throw new ArgumentNullException(nameof(changedClassicColorSlots));
            }
            A.ColorScheme? effectiveColors = theme.ColorScheme
                ?? source.SlideMasterPart?.ThemePart?.Theme?.ThemeElements?
                    .ColorScheme;
            return BuildMasterRecord(prototype, ReadColorScheme(effectiveColors),
                background: null,
                roundTripThemeRecords: BuildRoundTripThemeRecords(theme,
                    source.SlideLayout?.ColorMapOverride),
                rewriteColorScheme: changedClassicColorSlots.Count > 0,
                colorSchemeSlotsToRewrite: changedClassicColorSlots);
        }

        internal static byte[] BuildPreservedThemeRecord(
            LegacyPptRecord prototype, SlidePart source,
            IReadOnlyList<int> changedClassicColorSlots) {
            if (prototype == null) throw new ArgumentNullException(nameof(prototype));
            if (source == null) throw new ArgumentNullException(nameof(source));
            A.ThemeOverride theme = (source.ThemeOverridePart
                    ?? source.SlideLayoutPart?.ThemeOverridePart)?.ThemeOverride
                ?? throw new InvalidDataException(
                    "A projected binary slide has no DrawingML theme override to preserve.");
            if (changedClassicColorSlots == null) {
                throw new ArgumentNullException(nameof(changedClassicColorSlots));
            }
            A.ColorScheme? effectiveColors = theme.ColorScheme
                ?? source.SlideLayoutPart?.SlideMasterPart?.ThemePart?.Theme?
                    .ThemeElements?.ColorScheme;
            byte[] bytes = BuildMasterRecord(prototype,
                ReadColorScheme(effectiveColors), background: null,
                roundTripThemeRecords: BuildRoundTripThemeRecords(theme,
                    source.Slide?.ColorMapOverride
                        ?? source.SlideLayoutPart?.SlideLayout?
                            .ColorMapOverride),
                rewriteColorScheme: changedClassicColorSlots.Count > 0,
                colorSchemeSlotsToRewrite: changedClassicColorSlots);
            if (changedClassicColorSlots.Count == 0) return bytes;
            LegacyPptRecord record = LegacyPptRecordReader.ReadSingle(bytes, 0,
                new LegacyPptImportOptions());
            LegacyPptRecord atom = record.Children.First(child =>
                child.Type == RecordSlideAtom && child.PayloadLength >= 22);
            ushort flags = ReadUInt16(bytes, atom.PayloadOffset + 20);
            WriteUInt16(bytes, atom.PayloadOffset + 20,
                unchecked((ushort)(flags & ~0x0002)));
            return bytes;
        }

        internal static byte[] BuildPreservedThemeRecord(
            LegacyPptRecord prototype, NotesSlidePart source,
            IReadOnlyList<int> changedClassicColorSlots) {
            if (prototype == null) throw new ArgumentNullException(nameof(prototype));
            if (source == null) throw new ArgumentNullException(nameof(source));
            A.ThemeOverride theme = source.ThemeOverridePart?.ThemeOverride
                ?? throw new InvalidDataException(
                    "A projected binary notes page has no DrawingML theme override to preserve.");
            if (changedClassicColorSlots == null) {
                throw new ArgumentNullException(nameof(changedClassicColorSlots));
            }
            A.ColorScheme? effectiveColors = theme.ColorScheme
                ?? source.NotesMasterPart?.ThemePart?.Theme?.ThemeElements?
                    .ColorScheme;
            byte[] bytes = BuildMasterRecord(prototype,
                ReadColorScheme(effectiveColors), background: null,
                roundTripThemeRecords: BuildRoundTripThemeRecords(theme,
                    source.NotesSlide?.ColorMapOverride),
                rewriteColorScheme: changedClassicColorSlots.Count > 0,
                colorSchemeSlotsToRewrite: changedClassicColorSlots);
            if (changedClassicColorSlots.Count == 0) return bytes;
            LegacyPptRecord record = LegacyPptRecordReader.ReadSingle(bytes, 0,
                new LegacyPptImportOptions());
            LegacyPptRecord atom = record.Children.First(child =>
                child.Type == RecordNotesAtom && child.PayloadLength >= 6);
            ushort flags = ReadUInt16(bytes, atom.PayloadOffset + 4);
            WriteUInt16(bytes, atom.PayloadOffset + 4,
                unchecked((ushort)(flags & ~0x0002)));
            return bytes;
        }

        private static byte[] BuildPreservedMasterThemeRecord(
            LegacyPptRecord prototype, ThemePart? themePart,
            DocumentFormat.OpenXml.Presentation.ColorMap? colorMap,
            IReadOnlyList<int> changedClassicColorSlots) {
            A.Theme theme = themePart?.Theme
                ?? throw new InvalidDataException(
                    "A projected binary master has no DrawingML theme to preserve.");
            if (changedClassicColorSlots == null) {
                throw new ArgumentNullException(nameof(changedClassicColorSlots));
            }
            return BuildMasterRecord(prototype, ReadColorScheme(themePart),
                background: null,
                roundTripThemeRecords: BuildRoundTripThemeRecords(
                    theme, colorMap),
                rewriteColorScheme: changedClassicColorSlots.Count > 0,
                colorSchemeSlotsToRewrite: changedClassicColorSlots);
        }

        private static byte[] PatchColorSchemeAtom(LegacyPptRecord source,
            LegacyPptWriterColorScheme scheme,
            IReadOnlyList<int> slots) {
            if (source.PayloadLength < 32) {
                throw new InvalidDataException(
                    "The imported classic master color scheme is shorter than eight colors.");
            }
            byte[] record = source.CopyRecordBytes();
            foreach (int index in slots) {
                if (index < 0 || index >= scheme.Colors.Count) {
                    throw new InvalidDataException(
                        "A classic master color-scheme edit referenced an invalid slot.");
                }
                string color = scheme.Colors[index];
                int offset = 8 + index * 4;
                record[offset] = Convert.ToByte(color.Substring(0, 2), 16);
                record[offset + 1] = Convert.ToByte(color.Substring(2, 2), 16);
                record[offset + 2] = Convert.ToByte(color.Substring(4, 2), 16);
            }
            return record;
        }

        internal static byte[] BuildPreservedBackgroundRecord(
            LegacyPptRecord prototype,
            LegacyPptWriterBackground background,
            LegacyPptWriterPictureCatalog? pictureCatalog = null) {
            if (prototype == null) throw new ArgumentNullException(nameof(prototype));
            if (background == null) throw new ArgumentNullException(nameof(background));
            byte[] bytes = BuildMasterRecord(prototype,
                ReadColorScheme(themePart: null), background,
                rewriteColorScheme: false,
                pictureCatalog: pictureCatalog);
            LegacyPptRecord record = LegacyPptRecordReader.ReadSingle(bytes, 0,
                new LegacyPptImportOptions());
            LegacyPptRecord? slideAtom = record.Children.FirstOrDefault(child =>
                child.Type == RecordSlideAtom && child.PayloadLength >= 22);
            if (slideAtom != null) {
                ushort flags = ReadUInt16(bytes,
                    slideAtom.PayloadOffset + 20);
                WriteUInt16(bytes, slideAtom.PayloadOffset + 20,
                    unchecked((ushort)(flags & ~0x0004)));
            }
            LegacyPptRecord? notesAtom = record.Children.FirstOrDefault(child =>
                child.Type == RecordNotesAtom && child.PayloadLength >= 6);
            if (notesAtom != null) {
                ushort flags = ReadUInt16(bytes,
                    notesAtom.PayloadOffset + 4);
                WriteUInt16(bytes, notesAtom.PayloadOffset + 4,
                    unchecked((ushort)(flags & ~0x0004)));
            }
            return bytes;
        }

        internal static byte[] BuildPreservedMasterObjectInheritanceRecord(
            LegacyPptRecord prototype, bool followsMasterObjects) {
            if (prototype == null) throw new ArgumentNullException(nameof(prototype));
            byte[] bytes = prototype.CopyRecordBytes();
            LegacyPptRecord? slideAtom = prototype.Children.FirstOrDefault(child =>
                child.Type == RecordSlideAtom && child.PayloadLength >= 22);
            if (slideAtom != null) {
                int flagsOffset = checked(slideAtom.PayloadOffset
                    - prototype.Offset + 20);
                ushort flags = ReadUInt16(bytes, flagsOffset);
                flags = followsMasterObjects
                    ? unchecked((ushort)(flags | 0x0001))
                    : unchecked((ushort)(flags & ~0x0001));
                WriteUInt16(bytes, flagsOffset, flags);
                return bytes;
            }
            LegacyPptRecord? notesAtom = prototype.Children.FirstOrDefault(child =>
                child.Type == RecordNotesAtom && child.PayloadLength >= 6);
            if (notesAtom != null) {
                int flagsOffset = checked(notesAtom.PayloadOffset
                    - prototype.Offset + 4);
                ushort flags = ReadUInt16(bytes, flagsOffset);
                flags = followsMasterObjects
                    ? unchecked((ushort)(flags | 0x0001))
                    : unchecked((ushort)(flags & ~0x0001));
                WriteUInt16(bytes, flagsOffset, flags);
                return bytes;
            }
            throw new InvalidDataException(
                "The binary PowerPoint persist object has no slide or notes inheritance atom.");
        }

        internal static byte[] BuildPreservedPlaceholderSignatureRecord(
            LegacyPptRecord prototype, IReadOnlyList<PowerPointShape> shapes,
            LegacyPptWriterShapeContext shapeContext) {
            if (prototype == null) throw new ArgumentNullException(nameof(prototype));
            byte[] bytes = prototype.CopyRecordBytes();
            LegacyPptRecord? slideAtom = prototype.Children.FirstOrDefault(child =>
                child.Type == RecordSlideAtom && child.PayloadLength >= 12);
            if (slideAtom == null) return bytes;
            byte[] placeholders = BuildShapePlaceholderTypes(shapes,
                shapeContext);
            int placeholderOffset = checked(slideAtom.PayloadOffset
                - prototype.Offset + 4);
            Buffer.BlockCopy(placeholders, 0, bytes, placeholderOffset,
                placeholders.Length);
            return bytes;
        }

        private static byte[] BuildColorSchemeAtom(LegacyPptWriterColorScheme scheme) {
            var payload = new byte[32];
            for (int index = 0; index < scheme.Colors.Count; index++) {
                string color = scheme.Colors[index];
                payload[index * 4] = Convert.ToByte(color.Substring(0, 2), 16);
                payload[index * 4 + 1] = Convert.ToByte(color.Substring(2, 2), 16);
                payload[index * 4 + 2] = Convert.ToByte(color.Substring(4, 2), 16);
            }
            return BuildRecord(version: 0, instance: 1, RecordColorSchemeAtom, payload);
        }

        private static LegacyPptWriterColorScheme ReadColorScheme(ThemePart? themePart) {
            A.ColorScheme? source = themePart?.Theme?.ThemeElements?.ColorScheme;
            return ReadColorScheme(source);
        }

        private static LegacyPptWriterColorScheme ReadColorScheme(
            A.ColorScheme? source) {
            string Read(PowerPointThemeColor slot, string fallback) {
                OpenXmlCompositeElement? element = slot switch {
                    PowerPointThemeColor.Dark1 => source?.Dark1Color,
                    PowerPointThemeColor.Light1 => source?.Light1Color,
                    PowerPointThemeColor.Dark2 => source?.Dark2Color,
                    PowerPointThemeColor.Light2 => source?.Light2Color,
                    PowerPointThemeColor.Accent1 => source?.Accent1Color,
                    PowerPointThemeColor.Accent2 => source?.Accent2Color,
                    PowerPointThemeColor.Accent3 => source?.Accent3Color,
                    PowerPointThemeColor.Accent4 => source?.Accent4Color,
                    _ => null
                };
                string? value = element?
                    .GetFirstChild<DocumentFormat.OpenXml.Drawing.RgbColorModelHex>()?
                    .Val?.Value;
                if (string.IsNullOrWhiteSpace(value)) {
                    value = element?
                        .GetFirstChild<DocumentFormat.OpenXml.Drawing.SystemColor>()?
                        .LastColor?.Value;
                }
                return NormalizeColor(value, fallback);
            }

            return new LegacyPptWriterColorScheme(new[] {
                Read(PowerPointThemeColor.Light1, "FFFFFF"),
                Read(PowerPointThemeColor.Dark1, "000000"),
                Read(PowerPointThemeColor.Accent4, "808080"),
                Read(PowerPointThemeColor.Dark2, "000000"),
                Read(PowerPointThemeColor.Light2, "FFFFFF"),
                Read(PowerPointThemeColor.Accent1, "4472C4"),
                Read(PowerPointThemeColor.Accent2, "ED7D31"),
                Read(PowerPointThemeColor.Accent3, "A5A5A5")
            });
        }

        private static string NormalizeColor(string? value, string fallback) {
            string candidate = (value ?? string.Empty).Trim().TrimStart('#');
            return candidate.Length == 6 && candidate.All(Uri.IsHexDigit)
                ? candidate.ToUpperInvariant()
                : fallback;
        }

        private static byte[] BuildMasterList(int masterCount) {
            var children = new List<byte[]>(masterCount);
            for (int index = 0; index < masterCount; index++) {
                var payload = new byte[20];
                WriteUInt32(payload, 0,
                    checked(FirstMasterPersistId + unchecked((uint)index)));
                WriteUInt32(payload, 12,
                    checked(FirstMasterId + unchecked((uint)index)));
                children.Add(BuildRecord(version: 0, instance: 0,
                    RecordSlidePersistAtom, payload));
            }
            return BuildContainer(RecordSlideListWithText, instance: 1, children);
        }

        private sealed class LegacyPptWriterMasterCatalog {
            private readonly IReadOnlyDictionary<string, uint> _masterIds;

            internal LegacyPptWriterMasterCatalog(
                IReadOnlyDictionary<string, uint> masterIds,
                IReadOnlyList<byte[]> persistObjects,
                byte[] notesMasterPersistObject,
                byte[]? handoutMasterPersistObject, int count,
                IReadOnlyDictionary<uint, int> drawingShapeCounts,
                LegacyPptWriterFontCatalog fonts) {
                _masterIds = new ReadOnlyDictionary<string, uint>(
                    masterIds.ToDictionary(pair => pair.Key, pair => pair.Value,
                        StringComparer.Ordinal));
                PersistObjects = new ReadOnlyCollection<byte[]>(persistObjects.ToArray());
                NotesMasterPersistObject = notesMasterPersistObject;
                HandoutMasterPersistObject = handoutMasterPersistObject;
                Count = count;
                DrawingShapeCounts = new ReadOnlyDictionary<uint, int>(
                    drawingShapeCounts.ToDictionary(pair => pair.Key,
                        pair => pair.Value));
                Fonts = fonts;
            }

            internal int Count { get; }
            internal IReadOnlyList<byte[]> PersistObjects { get; }
            internal byte[] NotesMasterPersistObject { get; }
            internal byte[]? HandoutMasterPersistObject { get; }
            internal IReadOnlyDictionary<uint, int> DrawingShapeCounts { get; }
            internal LegacyPptWriterFontCatalog Fonts { get; }

            internal uint GetMasterId(PowerPointSlide slide) {
                SlideMasterPart? masterPart = slide.SlidePart.SlideLayoutPart?.SlideMasterPart;
                if (masterPart == null
                    || !_masterIds.TryGetValue(masterPart.Uri.ToString(), out uint masterId)) {
                    throw new InvalidDataException(
                        "A slide references a layout whose slide master is not part of the presentation.");
                }
                return masterId;
            }
        }

        private sealed class LegacyPptWriterColorScheme {
            internal LegacyPptWriterColorScheme(IReadOnlyList<string> colors) {
                if (colors.Count != 8) {
                    throw new ArgumentException(
                        "A classic binary PowerPoint color scheme contains eight colors.",
                        nameof(colors));
                }
                Colors = new ReadOnlyCollection<string>(colors.ToArray());
            }

            internal IReadOnlyList<string> Colors { get; }
        }
    }
}
