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

        private static IEnumerable<string?> ReadNumberingFontFamilies(LegacyDocElement numbering) =>
            numbering.IsNone
                ? Enumerable.Empty<string?>()
                : numbering.Descendants("rPr").Where(IsNumberingSymbolRunProperties).Select(properties => ReadSupportedRunFormatting(properties).FontFamily);

        /// <summary>Whether a <c>w:rPr</c> is a list level's numbering symbol run properties.</summary>
        private static bool IsNumberingSymbolRunProperties(LegacyDocElement properties) => properties.Parent.Is("lvl");

        private static void ValidateNativeNumberingReferences(LegacyDocSourcePackage source, IReadOnlyDictionary<string, ushort> styles) {
            LegacyDocElement numbering = source.Numbering;
            var instances = numbering.Elements("num")
                .Where(item => LegacyDocValues.Int32(item.Attr("numId")) > 0).ToDictionary(item => LegacyDocValues.Int32(item.Attr("numId"))!.Value);
            var definitions = numbering.Exists ? LegacyDocWordHelpers.GetCanonicalAbstractDefinitions(numbering)
                : new Dictionary<int, LegacyDocElement>();
            var roots = new List<LegacyDocElement> { source.Document, source.Footnotes,
                source.Endnotes, source.Comments };
            roots.AddRange(source.HeaderParts.Select(part => source.Root(part.PartName)));
            roots.AddRange(source.FooterParts.Select(part => source.Root(part.PartName)));
            roots.AddRange(source.Styles.Definitions
                .Where(style => LegacyDocStyleCatalog.StyleId(style) is string id && styles.ContainsKey(id)));
            foreach (LegacyDocElement properties in roots.Where(root => root.Exists).SelectMany(root => root.Descendants("numPr"))) {
                int? id = LegacyDocValues.Int32(properties.Child("numId").Val);
                if (!id.HasValue || id == 0) continue;
                int level = LegacyDocValues.Int32(properties.Child("ilvl").Val) ?? 0;
                if (!instances.TryGetValue(id.Value, out LegacyDocElement instance)
                    || !definitions.TryGetValue(LegacyDocValues.Int32(instance.Child("abstractNumId").Val) ?? -1, out LegacyDocElement definition)
                    || !definition.Elements("lvl").Any(item => LegacyDocValues.Int32(item.Attr("ilvl")) == level))
                    throw new NotSupportedException("Native DOC paragraph or style references an unavailable list instance or level.");
            }
        }

        private static LegacyDocWritableNumbering CreateWritableNumbering(LegacyDocElement numbering,
            IReadOnlyDictionary<string, ushort> styles, IReadOnlyDictionary<string, int> fonts) {
            var instances = numbering.Elements("num")
                .Where(item => LegacyDocValues.Int32(item.Attr("numId")) != 0).ToArray();
            if (instances.Length == 0) return new LegacyDocWritableNumbering(Array.Empty<byte>(), 0, Array.Empty<byte>());
            if (instances.Any(item => LegacyDocValues.Int32(item.Attr("numId")) is not int numberId || numberId < 1 || numberId > short.MaxValue)
                || instances.Select(item => LegacyDocValues.Int32(item.Attr("numId"))!.Value).Distinct().Count() != instances.Length)
                throw new NotSupportedException("Native DOC list instance ids must be unique values from 1 through 32767.");
            var referencedIds = new HashSet<int?>(instances.Select(instance => LegacyDocValues.Int32(instance.Child("abstractNumId").Val)));
            var referencedDefinitions = numbering.Elements("abstractNum")
                .Where(item => referencedIds.Contains(LegacyDocValues.Int32(item.Attr("abstractNumId")))).ToArray();
            if (referencedDefinitions.Length == 0 || referencedDefinitions.Length > short.MaxValue
                || referencedDefinitions.Any(item => LegacyDocValues.Int32(item.Attr("abstractNumId")) == null)
                || referencedDefinitions.Select(item => LegacyDocValues.Int32(item.Attr("abstractNumId"))!.Value).Distinct().Count() != referencedDefinitions.Length)
                throw new NotSupportedException("Native DOC lists require unique abstract numbering definitions.");
            Dictionary<int, LegacyDocElement> canonicalDefinitions = LegacyDocWordHelpers.GetCanonicalAbstractDefinitions(numbering);
            LegacyDocElement[] definitions = referencedDefinitions.Select(item => canonicalDefinitions[LegacyDocValues.Int32(item.Attr("abstractNumId"))!.Value]).Distinct().ToArray();
            var nativeByDefinition = definitions.Select((item, index) => new { definition = item, native = index + 1 })
                .ToDictionary(item => item.definition, item => item.native);
            // Different abstract IDs with one authored identity select the first definition
            // in Word. Preserve that shared sequence as one native lsid, including aliases.
            var nativeIds = canonicalDefinitions.Where(item => nativeByDefinition.ContainsKey(item.Value))
                .ToDictionary(item => item.Key, item => nativeByDefinition[item.Value]);
            using var listStream = new MemoryStream();
            WriteUInt16(listStream, checked((ushort)definitions.Length));
            var levelGroups = new List<LegacyDocElement[]>();
            foreach (LegacyDocElement definition in definitions) {
                string? restartAfterBreak = definition.AttrNs(Word2012Namespace, "restartNumberingAfterBreak");
                if (!string.IsNullOrEmpty(restartAfterBreak) && restartAfterBreak != "0" && restartAfterBreak != "false" && restartAfterBreak != "off")
                    throw new NotSupportedException("Native DOC saving does not support restarting a list after a section break.");
                if (definition.Child("numStyleLink").Exists || definition.Child("styleLink").Exists || definition.HasDescendant("lvlPicBulletId"))
                    throw new NotSupportedException("Native DOC saving does not support linked or picture-bullet list definitions.");
                LegacyDocElement[] levels = definition.Elements("lvl").OrderBy(item => LegacyDocValues.Int32(item.Attr("ilvl"))).ToArray();
                if (levels.Length != 1 && levels.Length != 9 || levels.Where((item, index) => LegacyDocValues.Int32(item.Attr("ilvl")) != index).Any())
                    throw new NotSupportedException("Native DOC lists need level 0 alone or all nine ordered levels.");
                levelGroups.Add(levels);
                WriteInt32(listStream, nativeIds[LegacyDocValues.Int32(definition.Attr("abstractNumId"))!.Value]);
                WriteInt32(listStream, 0);
                for (int index = 0; index < 9; index++) {
                    string? styleId = index < levels.Length ? levels[index].Child("pStyle").Val : null;
                    ushort styleIndex = NoBaseStyleIndex;
                    if (!string.IsNullOrEmpty(styleId) && !styles.TryGetValue(styleId!, out styleIndex)
                        && !TryMapBuiltInParagraphStyleIndex(styleId!, out styleIndex))
                        throw new NotSupportedException("Native DOC list level references an unavailable paragraph style.");
                    WriteUInt16(listStream, styleIndex);
                }
                listStream.WriteByte((byte)(levels.Length == 1 ? 1 : definition.Child("multiLevelType").Val == "hybridMultilevel" ? 16 : 0));
                listStream.WriteByte(0);
            }
            int headerLength = checked((int)listStream.Length);
            foreach (LegacyDocElement[] levels in levelGroups) foreach (LegacyDocElement level in levels) WriteListLevel(listStream, level, styles, fonts);

            // iLfo indexes this array directly. Unused sparse slots refer to a valid,
            // unreferenced definition and do not affect the authored instances.
            int maximumId = instances.Max(item => LegacyDocValues.Int32(item.Attr("numId"))!.Value);
            var byId = instances.ToDictionary(item => LegacyDocValues.Int32(item.Attr("numId"))!.Value);
            using var instanceStream = new MemoryStream();
            WriteInt32(instanceStream, maximumId);
            for (int id = 1; id <= maximumId; id++) {
                LegacyDocElement instance = byId.TryGetValue(id, out LegacyDocElement found) ? found : instances[0];
                if (!nativeIds.TryGetValue(LegacyDocValues.Int32(instance.Child("abstractNumId").Val) ?? -1, out int nativeId))
                    throw new NotSupportedException("Native DOC list instance references an unavailable definition.");
                LegacyDocElement[] overrides = byId.ContainsKey(id) ? instance.Elements("lvlOverride").ToArray() : Array.Empty<LegacyDocElement>();
                ValidateListOverrides(overrides, levelGroups[nativeId - 1].Length);
                WriteInt32(instanceStream, nativeId); WriteInt32(instanceStream, 0); WriteInt32(instanceStream, 0);
                instanceStream.WriteByte((byte)overrides.Length); instanceStream.WriteByte(0); WriteUInt16(instanceStream, 0);
            }
            for (int id = 1; id <= maximumId; id++) {
                WriteInt32(instanceStream, -1);
                if (!byId.TryGetValue(id, out LegacyDocElement instance)) continue;
                LegacyDocElement[] abstractLevels = levelGroups[nativeIds[LegacyDocValues.Int32(instance.Child("abstractNumId").Val)!.Value] - 1];
                foreach (LegacyDocElement item in instance.Elements("lvlOverride")) {
                    LegacyDocElement level = item.Child("lvl");
                    int? start = LegacyDocValues.Int32(item.Child("startOverride").Val);
                    if (start < 0 || start > short.MaxValue) throw new NotSupportedException("Native DOC list starts must be between 0 and 32767.");
                    WriteInt32(instanceStream, start ?? 0);
                    int levelIndex = LegacyDocValues.Int32(item.Attr("ilvl"))!.Value;
                    instanceStream.WriteByte((byte)(levelIndex | (start.HasValue ? 16 : 0) | (level.Exists ? 32 : 0)));
                    instanceStream.WriteByte(0); WriteUInt16(instanceStream, 0);
                    if (level.Exists) {
                        // The override level's symbol fonts resolve their theme selectors when they are read; resolving
                        // here keeps an unknown selector failing before the level's other checks.
                        LegacyDocElement levelFonts = level.Child("rPr").Child("rFonts");
                        if (levelFonts.Exists) _ = ResolveNativeThemeRunFonts(levelFonts);
                        // Word uses the abstract start unless w:startOverride explicitly
                        // requests a restart; a formatting-only w:lvl can carry another start.
                        int effectiveStart = start ?? LegacyDocValues.Int32(abstractLevels[levelIndex].Child("start").Val) ?? 1;
                        WriteListLevel(instanceStream, level, styles, fonts, effectiveStart);
                    }
                }
            }
            return new LegacyDocWritableNumbering(listStream.ToArray(), headerLength, instanceStream.ToArray());
        }

        private static void ValidateListOverrides(LegacyDocElement[] overrides, int levelCount) {
            if (overrides.Length > levelCount || overrides.Any(item => LegacyDocValues.Int32(item.Attr("ilvl")) is not int levelIndex || levelIndex < 0 || levelIndex >= levelCount)
                || overrides.Select(item => LegacyDocValues.Int32(item.Attr("ilvl"))!.Value).Distinct().Count() != overrides.Length)
                throw new NotSupportedException("Native DOC list overrides need unique levels from 0 through 8.");
            if (overrides.Any(item => item.Child("lvl") is { Exists: true } level && LegacyDocValues.Int32(level.Attr("ilvl")) != LegacyDocValues.Int32(item.Attr("ilvl"))!.Value))
                throw new NotSupportedException("Native DOC list override formatting must describe its own level.");
        }

        // startOverride: the start a list instance's level override uses in place of the level's own w:start.
        private static void WriteListLevel(Stream stream, LegacyDocElement level, IReadOnlyDictionary<string, ushort> styles, IReadOnlyDictionary<string, int> fonts, int? startOverride = null) {
            int index = LegacyDocValues.Int32(level.Attr("ilvl")) ?? -1, start = startOverride ?? LegacyDocValues.Int32(level.Child("start").Val) ?? 1;
            string format = level.Child("numFmt").Val ?? "decimal";
            byte? nfc = format == "bullet" ? (byte)23 : format == "none" ? (byte)255 : ToNfc(format);
            if (index < 0 || index > 8 || start < 0 || start > short.MaxValue || nfc == null || nfc == 8 || nfc == 9 || nfc == 15 || nfc == 19)
                throw new NotSupportedException("Native DOC list level has an unsupported format or start.");
            foreach (LegacyDocElement child in level.Elements()) {
                if (child.WName is not ("start" or "numFmt" or "lvlText" or "lvlJc" or "pPr" or "rPr" or "lvlRestart" or "suff" or "isLgl" or "pStyle"))
                    throw new NotSupportedException($"Native DOC saving does not support list level property '{child.LocalName}'.");
            }
            if (level.Child("rPr").Child("highlight").Exists)
                throw new NotSupportedException("Native DOC saving does not support highlight in numbering symbol properties.");
            string marker = level.Child("lvlText").Val ?? "%" + (index + 1) + ".";
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
            if (rawText.Length > ushort.MaxValue || format == "bullet" && (rawText.Length != 1 || placeholders.Count != 0))
                throw new NotSupportedException("Native DOC list marker text cannot be represented.");
            string alignment = LegacyDocWordHelpers.ReadLevelJustification(level.Child("lvlJc")) ?? "left";
            byte flags = alignment == "right" ? (byte)2 : alignment == "center" ? (byte)1
                : alignment == "left" ? (byte)0 : throw new NotSupportedException("Unsupported native list alignment.");
            LegacyDocElement legalNumbering = level.Child("isLgl");
            if (legalNumbering.Exists && IsOnOffEnabled(legalNumbering)) flags |= 4;
            int? restart = LegacyDocValues.Int32(level.Child("lvlRestart").Val);
            if (restart.HasValue) {
                if (restart.Value < 0 || restart.Value > index) throw new NotSupportedException("Unsupported native list restart level.");
                flags |= 8;
            }
            if (LegacyDocValues.OnOff(level.Attr("tentative")) == true) flags |= 128;
            string suffix = level.Child("suff").Val ?? "tab";
            byte follow = suffix == "tab" ? (byte)0 : suffix == "space" ? (byte)1
                : suffix == "nothing" ? (byte)2 : throw new NotSupportedException("Unsupported native list suffix.");
            var paragraphFormat = ReadSupportedParagraphFormattingCore(level.Child("pPr"), styles, allowParagraphStyleId: false, kind: LegacyDocParagraphPropertiesKind.ListLevel);
            byte[] upx = LegacyDocParagraphFormattingWriter.CreateStyleParagraphUpx(paragraphFormat);
            byte[] papx = upx.Length < 2 ? Array.Empty<byte>() : upx.Skip(2).ToArray();
            byte[] chpx = CreateCharacterGrpprl(ReadSupportedRunFormatting(level.Child("rPr")), fonts).ToArray();
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
