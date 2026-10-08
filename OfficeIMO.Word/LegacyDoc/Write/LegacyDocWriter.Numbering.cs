using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static void WriteInt32(Stream stream, int value) {
            var bytes = new byte[4];
            WriteInt32(bytes, 0, value);
            stream.Write(bytes, 0, bytes.Length);
        }

        private readonly struct LegacyDocWritableNumbering {
            internal LegacyDocWritableNumbering(byte[] definitions, int headerLength, byte[] instances) {
                Definitions = definitions; DefinitionHeaderLength = headerLength; Instances = instances;
            }
            internal byte[] Definitions { get; }
            internal int DefinitionHeaderLength { get; }
            internal byte[] Instances { get; }
        }

        private static IEnumerable<string?> ReadNumberingFontFamilies(Numbering? numbering) =>
            numbering?.Descendants<NumberingSymbolRunProperties>().Select(properties => ReadSupportedRunFormatting(properties).FontFamily)
            ?? Enumerable.Empty<string?>();

        private static void ValidateNativeNumberingReferences(MainDocumentPart mainPart, IReadOnlyDictionary<string, ushort> styles) {
            Numbering? numbering = mainPart.NumberingDefinitionsPart?.Numbering;
            var instances = (numbering?.Elements<NumberingInstance>() ?? Enumerable.Empty<NumberingInstance>())
                .Where(item => item.NumberID?.Value > 0).ToDictionary(item => item.NumberID!.Value);
            var definitions = numbering != null ? WordListNumberingResolver.GetCanonicalAbstractDefinitions(numbering)
                : new Dictionary<int, AbstractNum>();
            var roots = new List<OpenXmlElement?> { mainPart.Document, mainPart.FootnotesPart?.Footnotes,
                mainPart.EndnotesPart?.Endnotes, mainPart.WordprocessingCommentsPart?.Comments };
            roots.AddRange(mainPart.HeaderParts.Select(part => part.Header));
            roots.AddRange(mainPart.FooterParts.Select(part => part.Footer));
            roots.AddRange(mainPart.StyleDefinitionsPart?.Styles?.Elements<Style>()
                .Where(style => style.StyleId?.Value is string id && styles.ContainsKey(id)) ?? Enumerable.Empty<Style>());
            foreach (NumberingProperties properties in roots.Where(root => root != null).SelectMany(root => root!.Descendants<NumberingProperties>())) {
                int? id = properties.NumberingId?.Val?.Value;
                if (!id.HasValue || id == 0) continue;
                int level = properties.NumberingLevelReference?.Val?.Value ?? 0;
                if (!instances.TryGetValue(id.Value, out NumberingInstance? instance)
                    || !definitions.TryGetValue(instance.AbstractNumId?.Val?.Value ?? -1, out AbstractNum? definition)
                    || !definition.Elements<Level>().Any(item => item.LevelIndex?.Value == level))
                    throw new NotSupportedException("Native DOC paragraph or style references an unavailable list instance or level.");
            }
        }

        private static LegacyDocWritableNumbering CreateWritableNumbering(Numbering? numbering,
            IReadOnlyDictionary<string, ushort> styles, IReadOnlyDictionary<string, int> fonts) {
            var instances = (numbering?.Elements<NumberingInstance>() ?? Enumerable.Empty<NumberingInstance>())
                .Where(item => item.NumberID?.Value != 0).ToArray();
            if (instances.Length == 0) return new LegacyDocWritableNumbering(Array.Empty<byte>(), 0, Array.Empty<byte>());
            if (instances.Any(item => item.NumberID?.Value == null || item.NumberID.Value < 1 || item.NumberID.Value > short.MaxValue)
                || instances.Select(item => item.NumberID!.Value).Distinct().Count() != instances.Length)
                throw new NotSupportedException("Native DOC list instance ids must be unique values from 1 through 32767.");
            var referencedDefinitions = (numbering?.Elements<AbstractNum>() ?? Enumerable.Empty<AbstractNum>())
                .Where(item => instances.Any(instance => instance.AbstractNumId?.Val?.Value == item.AbstractNumberId?.Value)).ToArray();
            if (referencedDefinitions.Length == 0 || referencedDefinitions.Length > short.MaxValue
                || referencedDefinitions.Any(item => item.AbstractNumberId?.Value == null)
                || referencedDefinitions.Select(item => item.AbstractNumberId!.Value).Distinct().Count() != referencedDefinitions.Length)
                throw new NotSupportedException("Native DOC lists require unique abstract numbering definitions.");
            Dictionary<int, AbstractNum> canonicalDefinitions = WordListNumberingResolver.GetCanonicalAbstractDefinitions(numbering!);
            AbstractNum[] definitions = referencedDefinitions.Select(item => canonicalDefinitions[item.AbstractNumberId!.Value]).Distinct().ToArray();
            var nativeByDefinition = definitions.Select((item, index) => new { definition = item, native = index + 1 })
                .ToDictionary(item => item.definition, item => item.native);
            // Different abstract IDs with one authored identity select the first definition
            // in Word. Preserve that shared sequence as one native lsid, including aliases.
            var nativeIds = canonicalDefinitions.Where(item => nativeByDefinition.ContainsKey(item.Value))
                .ToDictionary(item => item.Key, item => nativeByDefinition[item.Value]);
            using var listStream = new MemoryStream();
            WriteUInt16(listStream, checked((ushort)definitions.Length));
            var levelGroups = new List<Level[]>();
            foreach (AbstractNum definition in definitions) {
                string? restartAfterBreak = definition.GetAttributes().FirstOrDefault(attribute =>
                    attribute.LocalName == "restartNumberingAfterBreak" && attribute.NamespaceUri == "http://schemas.microsoft.com/office/word/2012/wordml").Value;
                if (!string.IsNullOrEmpty(restartAfterBreak) && restartAfterBreak != "0" && restartAfterBreak != "false" && restartAfterBreak != "off")
                    throw new NotSupportedException("Native DOC saving does not support restarting a list after a section break.");
                if (definition.NumberingStyleLink != null || definition.StyleLink != null || definition.Descendants<LevelPictureBulletId>().Any())
                    throw new NotSupportedException("Native DOC saving does not support linked or picture-bullet list definitions.");
                Level[] levels = definition.Elements<Level>().OrderBy(item => item.LevelIndex?.Value).ToArray();
                if (levels.Length != 1 && levels.Length != 9 || levels.Where((item, index) => item.LevelIndex?.Value != index).Any())
                    throw new NotSupportedException("Native DOC lists need level 0 alone or all nine ordered levels.");
                levelGroups.Add(levels);
                WriteInt32(listStream, nativeIds[definition.AbstractNumberId!.Value]);
                WriteInt32(listStream, 0);
                for (int index = 0; index < 9; index++) {
                    string? styleId = index < levels.Length ? levels[index].ParagraphStyleIdInLevel?.Val?.Value : null;
                    ushort styleIndex = NoBaseStyleIndex;
                    if (!string.IsNullOrEmpty(styleId) && !styles.TryGetValue(styleId!, out styleIndex)
                        && !TryMapBuiltInParagraphStyleIndex(styleId!, out styleIndex))
                        throw new NotSupportedException("Native DOC list level references an unavailable paragraph style.");
                    WriteUInt16(listStream, styleIndex);
                }
                listStream.WriteByte((byte)(levels.Length == 1 ? 1 : definition.MultiLevelType?.Val?.Value == MultiLevelValues.HybridMultilevel ? 16 : 0));
                listStream.WriteByte(0);
            }
            int headerLength = checked((int)listStream.Length);
            foreach (Level[] levels in levelGroups) foreach (Level level in levels) WriteListLevel(listStream, level, styles, fonts);

            // iLfo indexes this array directly. Unused sparse slots refer to a valid,
            // unreferenced definition and do not affect the authored instances.
            int maximumId = instances.Max(item => item.NumberID!.Value);
            var byId = instances.ToDictionary(item => item.NumberID!.Value);
            using var instanceStream = new MemoryStream();
            WriteInt32(instanceStream, maximumId);
            for (int id = 1; id <= maximumId; id++) {
                NumberingInstance instance = byId.TryGetValue(id, out NumberingInstance? found) ? found : instances[0];
                if (!nativeIds.TryGetValue(instance.AbstractNumId?.Val?.Value ?? -1, out int nativeId))
                    throw new NotSupportedException("Native DOC list instance references an unavailable definition.");
                LevelOverride[] overrides = byId.ContainsKey(id) ? instance.Elements<LevelOverride>().ToArray() : Array.Empty<LevelOverride>();
                ValidateListOverrides(overrides, levelGroups[nativeId - 1].Length);
                WriteInt32(instanceStream, nativeId); WriteInt32(instanceStream, 0); WriteInt32(instanceStream, 0);
                instanceStream.WriteByte((byte)overrides.Length); instanceStream.WriteByte(0); WriteUInt16(instanceStream, 0);
            }
            for (int id = 1; id <= maximumId; id++) {
                WriteInt32(instanceStream, -1);
                if (!byId.TryGetValue(id, out NumberingInstance? instance)) continue;
                Level[] abstractLevels = levelGroups[nativeIds[instance.AbstractNumId!.Val!.Value] - 1];
                foreach (LevelOverride item in instance.Elements<LevelOverride>()) {
                    Level? level = item.GetFirstChild<Level>();
                    int? start = item.StartOverrideNumberingValue?.Val?.Value;
                    if (start < 0 || start > short.MaxValue) throw new NotSupportedException("Native DOC list starts must be between 0 and 32767.");
                    WriteInt32(instanceStream, start ?? 0);
                    instanceStream.WriteByte((byte)(item.LevelIndex!.Value | (start.HasValue ? 16 : 0) | (level != null ? 32 : 0)));
                    instanceStream.WriteByte(0); WriteUInt16(instanceStream, 0);
                    if (level != null) {
                        var effectiveLevel = (Level)level.CloneNode(true);
                        // Word uses the abstract start unless w:startOverride explicitly
                        // requests a restart; a formatting-only w:lvl can carry another start.
                        effectiveLevel.StartNumberingValue = new StartNumberingValue {
                            Val = start ?? abstractLevels[item.LevelIndex.Value].StartNumberingValue?.Val?.Value ?? 1
                        };
                        WriteListLevel(instanceStream, effectiveLevel, styles, fonts);
                    }
                }
            }
            return new LegacyDocWritableNumbering(listStream.ToArray(), headerLength, instanceStream.ToArray());
        }

        private static void ValidateListOverrides(LevelOverride[] overrides, int levelCount) {
            if (overrides.Length > levelCount || overrides.Any(item => item.LevelIndex?.Value == null || item.LevelIndex.Value < 0 || item.LevelIndex.Value >= levelCount)
                || overrides.Select(item => item.LevelIndex!.Value).Distinct().Count() != overrides.Length)
                throw new NotSupportedException("Native DOC list overrides need unique levels from 0 through 8.");
            if (overrides.Any(item => item.GetFirstChild<Level>() is Level level && level.LevelIndex?.Value != item.LevelIndex!.Value))
                throw new NotSupportedException("Native DOC list override formatting must describe its own level.");
        }

        private static void WriteListLevel(Stream stream, Level level, IReadOnlyDictionary<string, ushort> styles, IReadOnlyDictionary<string, int> fonts) {
            int index = level.LevelIndex?.Value ?? -1, start = level.StartNumberingValue?.Val?.Value ?? 1;
            NumberFormatValues format = level.NumberingFormat?.Val?.Value ?? NumberFormatValues.Decimal;
            byte? nfc = format == NumberFormatValues.Bullet ? (byte)23 : format == NumberFormatValues.None ? (byte)255 : LegacyDocNumberFormatMapper.ToNfc(format);
            if (index < 0 || index > 8 || start < 0 || start > short.MaxValue || nfc == null || nfc == 8 || nfc == 9 || nfc == 15 || nfc == 19)
                throw new NotSupportedException("Native DOC list level has an unsupported format or start.");
            foreach (OpenXmlElement child in level.ChildElements) {
                if (child is not StartNumberingValue && child is not NumberingFormat && child is not LevelText && child is not LevelJustification
                    && child is not PreviousParagraphProperties && child is not NumberingSymbolRunProperties && child is not LevelRestart
                    && child is not LevelSuffix && child is not IsLegalNumberingStyle && child is not ParagraphStyleIdInLevel)
                    throw new NotSupportedException($"Native DOC saving does not support list level property '{child.LocalName}'.");
            }
            string marker = level.LevelText?.Val?.Value ?? "%" + (index + 1) + ".";
            var rawText = new StringBuilder(marker.Length);
            var placeholders = new List<byte>();
            for (int position = 0; position < marker.Length; position++) {
                char character = marker[position];
                if (character == '%' && position + 1 < marker.Length && marker[position + 1] >= '1' && marker[position + 1] <= '9') {
                    int referenced = marker[++position] - '1';
                    if (referenced > index || rawText.Length >= byte.MaxValue || placeholders.Count > index)
                        throw new NotSupportedException("Native DOC list marker has an unsupported level placeholder.");
                    placeholders.Add((byte)(rawText.Length + 1)); rawText.Append((char)referenced);
                } else {
                    if (character < 32) throw new NotSupportedException("Native DOC list marker contains a control character.");
                    rawText.Append(character);
                }
            }
            if (rawText.Length > ushort.MaxValue || format == NumberFormatValues.Bullet && (rawText.Length != 1 || placeholders.Count != 0))
                throw new NotSupportedException("Native DOC list marker text cannot be represented.");
            LevelJustificationValues alignment = level.LevelJustification?.Val?.Value ?? LevelJustificationValues.Left;
            byte flags = alignment == LevelJustificationValues.Right ? (byte)2 : alignment == LevelJustificationValues.Center ? (byte)1
                : alignment == LevelJustificationValues.Left ? (byte)0 : throw new NotSupportedException("Unsupported native list alignment.");
            if (level.IsLegalNumberingStyle != null && IsOnOffEnabled(level.IsLegalNumberingStyle)) flags |= 4;
            int? restart = level.LevelRestart?.Val?.Value;
            if (restart.HasValue) {
                if (restart.Value < 0 || restart.Value > index) throw new NotSupportedException("Unsupported native list restart level.");
                flags |= 8;
            }
            if (level.Tentative?.Value == true) flags |= 128;
            LevelSuffixValues suffix = level.LevelSuffix?.Val?.Value ?? LevelSuffixValues.Tab;
            byte follow = suffix == LevelSuffixValues.Tab ? (byte)0 : suffix == LevelSuffixValues.Space ? (byte)1
                : suffix == LevelSuffixValues.Nothing ? (byte)2 : throw new NotSupportedException("Unsupported native list suffix.");
            var paragraphFormat = ReadSupportedParagraphFormattingCore(level.PreviousParagraphProperties, styles, allowParagraphStyleId: false);
            byte[] upx = LegacyDocParagraphFormattingWriter.CreateStyleParagraphUpx(paragraphFormat);
            byte[] papx = upx.Length < 2 ? Array.Empty<byte>() : upx.Skip(2).ToArray();
            byte[] chpx = CreateCharacterGrpprl(ReadSupportedRunFormatting(level.NumberingSymbolRunProperties), fonts).ToArray();
            if (papx.Length > byte.MaxValue || chpx.Length > byte.MaxValue) throw new NotSupportedException("Native DOC list level formatting is too large.");
            WriteInt32(stream, start); stream.WriteByte(nfc.Value); stream.WriteByte(flags);
            for (int position = 0; position < 9; position++) stream.WriteByte(position < placeholders.Count ? placeholders[position] : (byte)0);
            stream.WriteByte(follow); WriteInt32(stream, 0); WriteInt32(stream, 0);
            stream.WriteByte((byte)chpx.Length); stream.WriteByte((byte)papx.Length); stream.WriteByte((byte)(restart ?? 0)); stream.WriteByte(0);
            stream.Write(papx, 0, papx.Length); stream.Write(chpx, 0, chpx.Length);
            WriteUInt16(stream, checked((ushort)rawText.Length));
            byte[] text = Encoding.Unicode.GetBytes(rawText.ToString()); stream.Write(text, 0, text.Length);
        }
    }
}
