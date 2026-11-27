using System.Collections.ObjectModel;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        internal const int MaxNativeMasterCount = 4093;
        private const uint FirstMasterId = 0x80000000U;
        private const uint FirstMasterPersistId = 2U;
        private const ushort RecordHandoutForWrite = 0x0FC9;

        /// <summary>The presentation's slide masters in relationship order, as the SDK listed them.</summary>
        private static List<LegacyPptPartView> SlideMasterPartsInRelationshipOrder(
            LegacyPptPackageView view) =>
            view.Presentation.RelatedParts(LegacyPptPackageView.SlideMasterRelationshipType);

        private static LegacyPptWriterMasterCatalog ReadMasterCatalog(
            LegacyPptPackageView view,
            PowerPointDocument presentation,
            LegacyPptRecord templateDocument,
            IReadOnlyList<LegacyPptRecord> prototypes,
            LegacyPptRecord notesMasterPrototype,
            LegacyPptWriterTopology topology,
            LegacyPptWriterPictureBulletCatalog pictureBullets,
            LegacyPptWriterPictureCatalog pictureCatalog) {
            List<LegacyPptPartView> masterParts = SlideMasterPartsInRelationshipOrder(view);
            if (masterParts.Count == 0) {
                throw new InvalidDataException(
                    "The Open XML presentation has no slide master to encode.");
            }
            if (masterParts.Count > MaxNativeMasterCount) {
                throw new NotSupportedException(
                    $"The binary PowerPoint persist-directory format supports at most {MaxNativeMasterCount} slide masters; "
                    + $"the presentation contains {masterParts.Count}.");
            }
            if (prototypes.Count == 0) {
                throw new InvalidDataException(
                    "The embedded binary PowerPoint template has no slide-master prototype.");
            }
            if (!TryReadMasterTextStyles(view, presentation, templateDocument,
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
                if (index < masterParts.Count) {
                    LegacyPptPartView source = masterParts[index];
                    if (!TryReadBackground(presentation, source,
                            out LegacyPptWriterBackground? background,
                            out string? backgroundReason)) {
                        throw new NotSupportedException(backgroundReason);
                    }
                    IReadOnlyList<PowerPointShape> sourceShapes =
                        ReadMasterShapesForWrite(presentation, source, out _);
                    IReadOnlyList<PowerPointShape> supportedShapes = sourceShapes
                        .Where(IsSupportedMasterShape).ToArray();
                    uint drawingId = topology.GetMasterDrawingId(index);
                    IReadOnlyList<byte[]> roundTripThemeRecords =
                        BuildRoundTripThemeRecords(source.Theme.Root,
                            source.Root.Child(Pml, "clrMap"), isOverride: false);
                    persistObjects.Add(BuildMasterRecord(prototype,
                        ReadColorScheme(source.Theme), background,
                        supportedShapes, drawingId,
                        LegacyPptWriterShapeContext.MainMaster,
                        textStyles.Get(source),
                        textStyles.GetStyle9(source), roundTripThemeRecords,
                        fonts: textStyles.Fonts,
                        pictureBullets: pictureBullets,
                        pictureCatalog: pictureCatalog));
                    drawingShapeCounts.Add(drawingId,
                        CountDrawingShapes(supportedShapes, pictureCatalog));
                    masterIds.Add(masterParts[index].Name,
                        checked(FirstMasterId + unchecked((uint)index)));
                } else {
                    persistObjects.Add(prototype.CopyRecordBytes());
                }
            }
            LegacyPptPartView notesMasterPart = view.NotesMaster;
            LegacyPptWriterBackground? notesBackground = null;
            if (!notesMasterPart.IsNone
                && !TryReadBackground(presentation, notesMasterPart, out notesBackground,
                    out string? notesBackgroundReason)) {
                throw new NotSupportedException(notesBackgroundReason);
            }
            IReadOnlyList<PowerPointShape>? notesShapes = notesMasterPart.IsNone
                ? null
                : ReadMasterShapesForWrite(presentation, notesMasterPart, out _)
                    .Where(IsSupportedMasterShape).ToArray();
            uint notesDrawingId = topology.NotesMasterDrawingId;
            LegacyPptPartView notesTheme = !notesMasterPart.Theme.IsNone
                ? notesMasterPart.Theme
                : masterParts[0].Theme;
            byte[] notesMaster = BuildMasterRecord(notesMasterPrototype,
                ReadColorScheme(notesTheme), notesBackground,
                notesShapes, notesDrawingId,
                LegacyPptWriterShapeContext.NotesMaster,
                roundTripThemeRecords: BuildRoundTripThemeRecords(
                    notesTheme.Root,
                    notesMasterPart.Root.Child(Pml, "clrMap"), isOverride: false),
                fonts: textStyles.Fonts,
                pictureBullets: pictureBullets,
                pictureCatalog: pictureCatalog);
            if (notesShapes != null) {
                drawingShapeCounts.Add(notesDrawingId,
                    CountDrawingShapes(notesShapes, pictureCatalog));
            }
            LegacyPptPartView handoutMasterPart = view.HandoutMaster;
            byte[]? handoutMaster = null;
            if (!handoutMasterPart.IsNone) {
                if (!TryReadBackground(presentation, handoutMasterPart,
                        out LegacyPptWriterBackground? handoutBackground,
                        out string? handoutBackgroundReason)) {
                    throw new NotSupportedException(handoutBackgroundReason);
                }
                IReadOnlyList<PowerPointShape> handoutShapes =
                    ReadMasterShapesForWrite(presentation, handoutMasterPart, out _)
                        .Where(IsSupportedMasterShape).ToArray();
                LegacyPptPartView handoutTheme = !handoutMasterPart.Theme.IsNone
                    ? handoutMasterPart.Theme
                    : masterParts[0].Theme;
                handoutMaster = BuildHandoutMasterRecord(
                    notesMasterPrototype,
                    ReadColorScheme(handoutTheme), handoutBackground,
                    handoutShapes, topology.HandoutMasterDrawingId,
                    BuildRoundTripThemeRecords(
                        handoutTheme.Root,
                        handoutMasterPart.Root.Child(Pml, "clrMap"), isOverride: false),
                    textStyles.Fonts, pictureBullets, pictureCatalog);
                drawingShapeCounts.Add(topology.HandoutMasterDrawingId,
                    CountDrawingShapes(handoutShapes, pictureCatalog));
            }
            return new LegacyPptWriterMasterCatalog(masterIds, persistObjects,
                notesMaster, handoutMaster, masterParts.Count,
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
        /// Rewrites only the classic color scheme and DrawingML round-trip theme records of an imported master: a slide,
        /// notes or handout master, or a title master's layout (its theme override). Every unrelated source child record
        /// is retained byte-for-byte.
        /// </summary>
        internal static byte[] BuildPreservedMasterThemeRecord(
            LegacyPptRecord prototype, LegacyPptPartView source,
            IReadOnlyList<int> changedClassicColorSlots) {
            ArgumentNullException.ThrowIfNull(prototype);
            if (source.IsNone) throw new ArgumentNullException(nameof(source));
            if (source.IsSlideLayout) {
                OfficeXmlNode theme = source.ThemeOverride.Root;
                if (theme.IsNone) {
                    throw new InvalidDataException(
                        "A projected binary title master has no DrawingML theme override to preserve.");
                }
                ArgumentNullException.ThrowIfNull(changedClassicColorSlots);
                OfficeXmlNode effectiveColors = theme.Child(Dml, "clrScheme");
                if (effectiveColors.IsNone) {
                    effectiveColors = ThemeElement(source.SlideMaster.Theme, "clrScheme");
                }
                return BuildMasterRecord(prototype, ReadColorScheme(effectiveColors),
                    background: null,
                    roundTripThemeRecords: BuildRoundTripThemeRecords(theme,
                        source.Root.Child(Pml, "clrMapOvr"), isOverride: true),
                    rewriteColorScheme: changedClassicColorSlots.Count > 0,
                    colorSchemeSlotsToRewrite: changedClassicColorSlots);
            }
            LegacyPptPartView themePart = source.Theme;
            OfficeXmlNode masterTheme = themePart.Root;
            if (masterTheme.IsNone) {
                throw new InvalidDataException(
                    "A projected binary master has no DrawingML theme to preserve.");
            }
            ArgumentNullException.ThrowIfNull(changedClassicColorSlots);
            return BuildMasterRecord(prototype, ReadColorScheme(themePart),
                background: null,
                roundTripThemeRecords: BuildRoundTripThemeRecords(
                    masterTheme, source.Root.Child(Pml, "clrMap"), isOverride: false),
                rewriteColorScheme: changedClassicColorSlots.Count > 0,
                colorSchemeSlotsToRewrite: changedClassicColorSlots);
        }

        /// <summary>
        /// Rewrites only the classic color scheme and DrawingML round-trip theme override records of an imported slide
        /// or notes page.
        /// </summary>
        internal static byte[] BuildPreservedThemeRecord(
            LegacyPptRecord prototype, LegacyPptPartView source,
            IReadOnlyList<int> changedClassicColorSlots) {
            ArgumentNullException.ThrowIfNull(prototype);
            if (source.IsNone) throw new ArgumentNullException(nameof(source));
            return source.IsNotesSlide
                ? BuildPreservedNotesThemeRecord(prototype, source, changedClassicColorSlots)
                : BuildPreservedSlideThemeRecord(prototype, source, changedClassicColorSlots);
        }

        private static byte[] BuildPreservedSlideThemeRecord(
            LegacyPptRecord prototype, LegacyPptPartView source,
            IReadOnlyList<int> changedClassicColorSlots) {
            LegacyPptPartView layout = source.SlideLayout;
            LegacyPptPartView themePart = !source.ThemeOverride.IsNone
                ? source.ThemeOverride
                : layout.ThemeOverride;
            OfficeXmlNode theme = themePart.Root;
            if (theme.IsNone) {
                throw new InvalidDataException(
                    "A projected binary slide has no DrawingML theme override to preserve.");
            }
            ArgumentNullException.ThrowIfNull(changedClassicColorSlots);
            OfficeXmlNode effectiveColors = theme.Child(Dml, "clrScheme");
            if (effectiveColors.IsNone) {
                effectiveColors = ThemeElement(layout.SlideMaster.Theme, "clrScheme");
            }
            OfficeXmlNode colorMap = source.Root.Child(Pml, "clrMapOvr");
            if (colorMap.IsNone) colorMap = layout.Root.Child(Pml, "clrMapOvr");
            byte[] bytes = BuildMasterRecord(prototype,
                ReadColorScheme(effectiveColors), background: null,
                roundTripThemeRecords: BuildRoundTripThemeRecords(theme,
                    colorMap, isOverride: true),
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

        private static byte[] BuildPreservedNotesThemeRecord(
            LegacyPptRecord prototype, LegacyPptPartView source,
            IReadOnlyList<int> changedClassicColorSlots) {
            OfficeXmlNode theme = source.ThemeOverride.Root;
            if (theme.IsNone) {
                throw new InvalidDataException(
                    "A projected binary notes page has no DrawingML theme override to preserve.");
            }
            ArgumentNullException.ThrowIfNull(changedClassicColorSlots);
            OfficeXmlNode effectiveColors = theme.Child(Dml, "clrScheme");
            if (effectiveColors.IsNone) {
                effectiveColors = ThemeElement(
                    source.RelatedPart(LegacyPptPackageView.NotesMasterRelationshipType).Theme,
                    "clrScheme");
            }
            byte[] bytes = BuildMasterRecord(prototype,
                ReadColorScheme(effectiveColors), background: null,
                roundTripThemeRecords: BuildRoundTripThemeRecords(theme,
                    source.Root.Child(Pml, "clrMapOvr"), isOverride: true),
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
            ArgumentNullException.ThrowIfNull(prototype);
            ArgumentNullException.ThrowIfNull(background);
            byte[] bytes = BuildMasterRecord(prototype,
                ReadColorScheme(default(OfficeXmlNode)), background,
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
            ArgumentNullException.ThrowIfNull(prototype);
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
            ArgumentNullException.ThrowIfNull(prototype);
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

        private static LegacyPptWriterColorScheme ReadColorScheme(LegacyPptPartView themePart) =>
            ReadColorScheme(ThemeElement(themePart, "clrScheme"));

        private static LegacyPptWriterColorScheme ReadColorScheme(
            OfficeXmlNode source) {
            string Read(string slot, string fallback) {
                OfficeXmlNode element = source.Child(Dml, slot);
                string? value = element.Child(Dml, "srgbClr").Attribute("val");
                if (string.IsNullOrWhiteSpace(value)) {
                    value = element.Child(Dml, "sysClr").Attribute("lastClr");
                }
                return NormalizeColor(value, fallback);
            }

            return new LegacyPptWriterColorScheme(new[] {
                Read("lt1", "FFFFFF"),
                Read("dk1", "000000"),
                Read("accent4", "808080"),
                Read("dk2", "000000"),
                Read("lt2", "FFFFFF"),
                Read("accent1", "4472C4"),
                Read("accent2", "ED7D31"),
                Read("accent3", "A5A5A5")
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

            internal uint GetMasterId(LegacyPptPackageView view, PowerPointSlide slide) {
                LegacyPptPartView masterPart = SlidePartView(view, slide).SlideLayout.SlideMaster;
                if (masterPart.IsNone
                    || !_masterIds.TryGetValue(masterPart.Name, out uint masterId)) {
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
