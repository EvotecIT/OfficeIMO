using System.Collections.ObjectModel;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        private const ushort RecordFontCollectionForWrite = 0x07D5;
        private const ushort RecordTextMasterStyleAtomForWrite = 0x0FA3;
        private const ushort RecordTextMasterStyle9AtomForWrite = 0x0FAD;
        private const ushort RecordFontEntityAtomForWrite = 0x0FB7;

        internal static bool CanWriteMasterTextStyles(
            PowerPointPresentation presentation, out string? reason) {
            if (!TryReadPictureBulletCatalog(presentation,
                    out LegacyPptWriterPictureBulletCatalog pictureBullets,
                    out reason)) return false;
            return TryReadMasterTextStyles(presentation,
                Template.Value.Document, pictureBullets, out _, out reason);
        }

        private static bool TryReadMasterTextStyles(
            PowerPointPresentation presentation, LegacyPptRecord templateDocument,
            LegacyPptWriterPictureBulletCatalog pictureBullets,
            out LegacyPptWriterMasterTextStyleCatalog catalog,
            out string? reason) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            var fonts = new LegacyPptWriterFontCatalog(templateDocument);
            var records = new Dictionary<string, IReadOnlyList<byte[]>>(
                StringComparer.Ordinal);
            var style9Records = new Dictionary<string,
                IReadOnlyList<byte[]>>(StringComparer.Ordinal);
            foreach (SlideMasterPart masterPart in presentation.OpenXmlDocument
                         .PresentationPart?.SlideMasterParts
                     ?? Enumerable.Empty<SlideMasterPart>()) {
                if (!TryBuildMasterTextStyleRecords(
                        masterPart.SlideMaster?.TextStyles, fonts,
                        pictureBullets,
                        out IReadOnlyList<byte[]> masterRecords,
                        out IReadOnlyList<byte[]> masterStyle9Records,
                        out reason)) {
                    catalog = new LegacyPptWriterMasterTextStyleCatalog(
                        records, style9Records, fonts);
                    return false;
                }
                records.Add(masterPart.Uri.ToString(), masterRecords);
                style9Records.Add(masterPart.Uri.ToString(),
                    masterStyle9Records);
            }
            catalog = new LegacyPptWriterMasterTextStyleCatalog(records,
                style9Records, fonts);
            reason = null;
            return true;
        }

        internal static bool TryBuildMasterTextStyleRecords(P.TextStyles? styles,
            LegacyPptWriterFontCatalog fonts, out IReadOnlyList<byte[]> records,
            out IReadOnlyList<byte[]> style9Records,
            out string? reason) => TryBuildMasterTextStyleRecords(styles,
                fonts, LegacyPptWriterPictureBulletCatalog.Empty,
                out records, out style9Records, out reason);

        internal static bool TryBuildMasterTextStyleRecords(P.TextStyles? styles,
            LegacyPptWriterFontCatalog fonts,
            LegacyPptWriterPictureBulletCatalog pictureBullets,
            out IReadOnlyList<byte[]> records,
            out IReadOnlyList<byte[]> style9Records,
            out string? reason) {
            var result = new List<byte[]>(3);
            var style9Result = new List<byte[]>(3);
            records = result;
            style9Records = style9Result;
            reason = null;
            if (styles == null) return true;
            if (!TryBuildMasterTextStyleRecord(styles.TitleStyle,
                    instance: (ushort)LegacyPptTextType.Title, fonts,
                    out byte[] title, out reason)
                || !TryBuildMasterTextStyleRecord(styles.BodyStyle,
                    instance: (ushort)LegacyPptTextType.Body, fonts,
                    out byte[] body, out reason)
                || !TryBuildMasterTextStyleRecord(styles.OtherStyle,
                    instance: (ushort)LegacyPptTextType.Other, fonts,
                    out byte[] other, out reason)) {
                return false;
            }
            result.Add(title);
            result.Add(body);
            result.Add(other);
            if (!TryBuildMasterTextStyle9Record(styles.TitleStyle,
                    instance: (ushort)LegacyPptTextType.Title,
                    pictureBullets,
                    out byte[]? title9, out reason)
                || !TryBuildMasterTextStyle9Record(styles.BodyStyle,
                    instance: (ushort)LegacyPptTextType.Body,
                    pictureBullets,
                    out byte[]? body9, out reason)
                || !TryBuildMasterTextStyle9Record(styles.OtherStyle,
                    instance: (ushort)LegacyPptTextType.Other,
                    pictureBullets,
                    out byte[]? other9, out reason)) return false;
            if (title9 != null) style9Result.Add(title9);
            if (body9 != null) style9Result.Add(body9);
            if (other9 != null) style9Result.Add(other9);
            return true;
        }

        internal static bool TryRewriteMasterTextStyleRecords(
            LegacyPptRecord master, P.TextStyles? styles,
            LegacyPptWriterFontCatalog fonts, out byte[] bytes,
            out string? reason) => TryRewriteMasterTextStyleRecords(master,
                styles, fonts, LegacyPptWriterPictureBulletCatalog.Empty,
                out bytes, out reason);

        internal static bool TryRewriteMasterTextStyleRecords(
            LegacyPptRecord master, P.TextStyles? styles,
            LegacyPptWriterFontCatalog fonts,
            LegacyPptWriterPictureBulletCatalog pictureBullets,
            out byte[] bytes,
            out string? reason) {
            if (master == null) throw new ArgumentNullException(nameof(master));
            if (fonts == null) throw new ArgumentNullException(nameof(fonts));
            bytes = master.CopyRecordBytes();
            reason = null;
            if (master.Version != 0x0F
                || !TryBuildMasterTextStyleRecords(styles, fonts,
                    pictureBullets,
                    out IReadOnlyList<byte[]> records,
                    out IReadOnlyList<byte[]> style9Records, out reason)) {
                return false;
            }
            var children = new List<byte[]>(master.Children.Count
                + records.Count);
            var replacements = records.ToDictionary(record =>
                LegacyPptRecordReader.ReadSingle(record, 0,
                    new LegacyPptImportOptions()).Instance);
            foreach (LegacyPptRecord child in master.Children) {
                if (child.Type == RecordTextMasterStyleAtomForWrite
                    && replacements.TryGetValue(child.Instance,
                        out byte[]? replacement)) {
                    children.Add(replacement);
                    replacements.Remove(child.Instance);
                    continue;
                }
                children.Add(child.CopyRecordBytes());
            }
            foreach (byte[] record in records) {
                ushort instance = LegacyPptRecordReader.ReadSingle(record, 0,
                    new LegacyPptImportOptions()).Instance;
                if (replacements.ContainsKey(instance)) {
                    children.Add(record);
                }
            }
            byte[] baseBytes = BuildContainer(master.Type, master.Instance,
                children);
            LegacyPptRecord rewritten = LegacyPptRecordReader.ReadSingle(
                baseBytes, 0, new LegacyPptImportOptions());
            if (!TryRewriteMasterTextStyle9Records(rewritten, style9Records,
                    new ushort[] {
                        (ushort)LegacyPptTextType.Title,
                        (ushort)LegacyPptTextType.Body,
                        (ushort)LegacyPptTextType.Other
                    }, replaceAllExisting: false, out bytes)) {
                reason = "The binary master contains malformed or duplicate PPT9 programmable tags.";
                return false;
            }
            reason = null;
            return true;
        }

        internal static bool TryRewriteMasterTextStyle9Records(
            LegacyPptRecord master, IReadOnlyList<byte[]> style9Records,
            IReadOnlyCollection<ushort>? instancesToReplace,
            bool replaceAllExisting, out byte[] bytes) {
            if (master == null) throw new ArgumentNullException(nameof(master));
            if (style9Records == null) {
                throw new ArgumentNullException(nameof(style9Records));
            }
            if (!replaceAllExisting && instancesToReplace == null) {
                throw new ArgumentNullException(nameof(instancesToReplace));
            }
            bytes = master.CopyRecordBytes();
            if (master.Version != 0x0F) return false;
            LegacyPptRecord[] progTags = master.Children.Where(child =>
                child.Type == RecordProgTags).ToArray();
            LegacyPptRecord[] ppt9Tags = progTags.SelectMany(tags =>
                tags.Children.Where(IsPpt9BinaryTag)).ToArray();
            if (ppt9Tags.Length > 1) return false;

            byte[]? rewrittenPpt9 = null;
            LegacyPptRecord? owner = null;
            if (ppt9Tags.Length == 1) {
                LegacyPptRecord existing = ppt9Tags[0];
                owner = progTags.Single(tags => tags.Children.Any(child =>
                    ReferenceEquals(child, existing)));
                if (owner.Version != 0x0F || owner.Instance != 0) {
                    return false;
                }
                if (!TryRewriteMasterTextStyle9BinaryTag(existing,
                        style9Records, instancesToReplace,
                        replaceAllExisting, out rewrittenPpt9)) return false;
            } else if (style9Records.Count > 0) {
                rewrittenPpt9 = BuildMasterTextStyle9BinaryTag(
                    style9Records);
                owner = progTags.FirstOrDefault(tags =>
                    tags.Version == 0x0F && tags.Instance == 0);
            }

            byte[]? rewrittenOwner = null;
            if (owner != null) {
                var tagChildren = new List<byte[]>(owner.Children.Count + 1);
                foreach (LegacyPptRecord child in owner.Children) {
                    if (ppt9Tags.Length == 1
                        && ReferenceEquals(child, ppt9Tags[0])) {
                        if (rewrittenPpt9 != null) {
                            tagChildren.Add(rewrittenPpt9);
                        }
                    } else {
                        tagChildren.Add(child.CopyRecordBytes());
                    }
                }
                if (ppt9Tags.Length == 0 && rewrittenPpt9 != null) {
                    tagChildren.Add(rewrittenPpt9);
                }
                rewrittenOwner = BuildRecord(owner.Version, owner.Instance,
                    owner.Type, Concat(tagChildren));
            }

            var masterChildren = new List<byte[]>(master.Children.Count + 1);
            foreach (LegacyPptRecord child in master.Children) {
                masterChildren.Add(owner != null && ReferenceEquals(child, owner)
                    ? rewrittenOwner!
                    : child.CopyRecordBytes());
            }
            if (owner == null && rewrittenPpt9 != null) {
                masterChildren.Add(BuildContainer(RecordProgTags, instance: 0,
                    new[] { rewrittenPpt9 }));
            }
            bytes = BuildRecord(master.Version, master.Instance, master.Type,
                Concat(masterChildren));
            return true;
        }

        private static bool TryRewriteMasterTextStyle9BinaryTag(
            LegacyPptRecord binaryTag,
            IReadOnlyList<byte[]> style9Records,
            IReadOnlyCollection<ushort>? instancesToReplace,
            bool replaceAllExisting, out byte[]? bytes) {
            bytes = binaryTag.CopyRecordBytes();
            LegacyPptRecord[] blobs = binaryTag.Children.Where(child =>
                child.Type == RecordBinaryTagDataBlob).ToArray();
            if (binaryTag.Version != 0x0F || binaryTag.Instance != 0
                || blobs.Length != 1 || blobs[0].Version != 0
                || blobs[0].Instance != 0) return false;
            IReadOnlyList<LegacyPptRecord> dataRecords;
            try {
                dataRecords = LegacyPptRecordReader.ReadSequence(
                    blobs[0].CopyRecordBytes(), 8, blobs[0].PayloadLength,
                    new LegacyPptImportOptions());
            } catch (Exception exception) when (exception
                is InvalidDataException or OverflowException
                    or ArgumentOutOfRangeException) {
                return false;
            }
            var rewrittenData = new List<byte[]>(style9Records.Count
                + dataRecords.Count);
            rewrittenData.AddRange(style9Records);
            rewrittenData.AddRange(dataRecords.Where(record => record.Type
                    != RecordTextMasterStyle9AtomForWrite
                    || !replaceAllExisting && !instancesToReplace!
                        .Contains(record.Instance))
                .Select(record => record.CopyRecordBytes()));
            if (rewrittenData.Count == 0) {
                bytes = null;
                return true;
            }
            byte[] blob = BuildRecord(blobs[0].Version, blobs[0].Instance,
                blobs[0].Type, Concat(rewrittenData));
            var children = new List<byte[]>(binaryTag.Children.Count);
            foreach (LegacyPptRecord child in binaryTag.Children) {
                children.Add(ReferenceEquals(child, blobs[0])
                    ? blob
                    : child.CopyRecordBytes());
            }
            bytes = BuildRecord(binaryTag.Version, binaryTag.Instance,
                binaryTag.Type, Concat(children));
            return true;
        }

        private static byte[] BuildMasterTextStyle9BinaryTag(
            IReadOnlyList<byte[]> style9Records) {
            byte[] tagName = BuildRecord(version: 0, instance: 0,
                RecordCString, Encoding.Unicode.GetBytes(Ppt9TagName));
            byte[] data = BuildRecord(version: 0, instance: 0,
                RecordBinaryTagDataBlob, Concat(style9Records));
            return BuildContainer(RecordProgBinaryTag, instance: 0,
                new[] { tagName, data });
        }

        private static bool TryBuildMasterTextStyleRecord(
            OpenXmlCompositeElement? style, ushort instance,
            LegacyPptWriterFontCatalog fonts, out byte[] record,
            out string? reason) {
            record = Array.Empty<byte>();
            reason = null;
            if (style == null) {
                record = BuildRecord(version: 0, instance,
                    RecordTextMasterStyleAtomForWrite, new byte[2]);
                return true;
            }
            if (style.HasAttributes) {
                reason = $"The {style.LocalName} master text style has attributes that are not represented by the base binary style atom.";
                return false;
            }
            if (!TryReadMasterTextStyleLevels(style, out Dictionary<int,
                    A.TextParagraphPropertiesType> levels,
                    out reason)) return false;
            int levelCount = levels.Count == 0 ? 0 : levels.Keys.Max() + 1;
            if (levelCount > 5) {
                reason = "Base binary PowerPoint master text styles support at most five levels.";
                return false;
            }
            using var payload = new MemoryStream();
            WriteUInt16(payload, checked((ushort)levelCount));
            for (int level = 0; level < levelCount; level++) {
                levels.TryGetValue(level,
                    out A.TextParagraphPropertiesType? properties);
                if (!TryWriteParagraphException(payload, properties, fonts,
                        out reason, allowAutoNumbering: true)
                    || !TryWriteCharacterException(payload,
                        properties?.GetFirstChild<A.DefaultRunProperties>(),
                        fonts, out reason)) {
                    return false;
                }
            }
            record = BuildRecord(version: 0, instance,
                RecordTextMasterStyleAtomForWrite, payload.ToArray());
            return true;
        }

        private static bool TryBuildMasterTextStyle9Record(
            OpenXmlCompositeElement? style, ushort instance,
            LegacyPptWriterPictureBulletCatalog pictureBullets,
            out byte[]? record, out string? reason) {
            record = null;
            reason = null;
            if (style == null) return true;
            if (!TryReadMasterTextStyleLevels(style, out Dictionary<int,
                    A.TextParagraphPropertiesType> levels,
                    out reason)) return false;
            if (!levels.Values.Any(properties => properties.ChildElements
                    .Any(child => child is A.AutoNumberedBullet
                        or A.PictureBullet))) {
                return true;
            }
            int levelCount = levels.Keys.Max() + 1;
            using var payload = new MemoryStream();
            WriteUInt16(payload, checked((ushort)levelCount));
            for (int level = 0; level < levelCount; level++) {
                levels.TryGetValue(level,
                    out A.TextParagraphPropertiesType? properties);
                A.AutoNumberedBullet? numbering = properties?
                    .GetFirstChild<A.AutoNumberedBullet>();
                if (!TryWriteAutomaticNumberingException9(payload,
                        numbering, numbering?.StartAt?.Value ?? 1,
                        properties?.GetFirstChild<A.PictureBullet>(),
                        pictureBullets,
                        out reason)) return false;
                WriteUInt32(payload, 0);
            }
            record = BuildRecord(version: 0, instance,
                RecordTextMasterStyle9AtomForWrite, payload.ToArray());
            return true;
        }

        private static bool TryReadMasterTextStyleLevels(
            OpenXmlCompositeElement style,
            out Dictionary<int, A.TextParagraphPropertiesType> levels,
            out string? reason) {
            levels = new Dictionary<int, A.TextParagraphPropertiesType>();
            reason = null;
            if (style.HasAttributes) {
                reason = $"The {style.LocalName} master text style has attributes that are not represented by the base binary style atom.";
                return false;
            }
            foreach (OpenXmlElement child in style.ChildElements) {
                if (child is not A.TextParagraphPropertiesType properties
                    || !TryGetMasterTextStyleLevel(properties, out int level)
                    || levels.ContainsKey(level)) {
                    reason = $"The {style.LocalName} master text style contains unsupported or duplicate level content.";
                    return false;
                }
                levels.Add(level, properties);
            }
            return true;
        }

        private static bool TryGetMasterTextStyleLevel(
            A.TextParagraphPropertiesType properties, out int level) {
            level = properties switch {
                A.Level1ParagraphProperties => 0,
                A.Level2ParagraphProperties => 1,
                A.Level3ParagraphProperties => 2,
                A.Level4ParagraphProperties => 3,
                A.Level5ParagraphProperties => 4,
                _ => -1
            };
            return level >= 0;
        }

        private static void WriteUInt16(Stream stream, ushort value) {
            stream.WriteByte(unchecked((byte)value));
            stream.WriteByte(unchecked((byte)(value >> 8)));
        }

        private static void WriteInt16(Stream stream, short value) =>
            WriteUInt16(stream, unchecked((ushort)value));

        private static void WriteUInt32(Stream stream, uint value) {
            stream.WriteByte(unchecked((byte)value));
            stream.WriteByte(unchecked((byte)(value >> 8)));
            stream.WriteByte(unchecked((byte)(value >> 16)));
            stream.WriteByte(unchecked((byte)(value >> 24)));
        }

        private sealed class LegacyPptWriterMasterTextStyleCatalog {
            private readonly IReadOnlyDictionary<string, IReadOnlyList<byte[]>>
                _records;
            private readonly IReadOnlyDictionary<string, IReadOnlyList<byte[]>>
                _style9Records;

            internal LegacyPptWriterMasterTextStyleCatalog(
                IReadOnlyDictionary<string, IReadOnlyList<byte[]>> records,
                IReadOnlyDictionary<string, IReadOnlyList<byte[]>>
                    style9Records,
                LegacyPptWriterFontCatalog fonts) {
                _records = new ReadOnlyDictionary<string,
                    IReadOnlyList<byte[]>>(records.ToDictionary(pair => pair.Key,
                    pair => (IReadOnlyList<byte[]>)new ReadOnlyCollection<byte[]>(
                        pair.Value.Select(record => record.ToArray()).ToArray()),
                    StringComparer.Ordinal));
                _style9Records = new ReadOnlyDictionary<string,
                    IReadOnlyList<byte[]>>(style9Records.ToDictionary(pair =>
                        pair.Key,
                        pair => (IReadOnlyList<byte[]>)
                            new ReadOnlyCollection<byte[]>(pair.Value.Select(
                                record => record.ToArray()).ToArray()),
                        StringComparer.Ordinal));
                Fonts = fonts;
            }

            internal LegacyPptWriterFontCatalog Fonts { get; }

            internal IReadOnlyList<byte[]> Get(SlideMasterPart masterPart) =>
                _records.TryGetValue(masterPart.Uri.ToString(),
                    out IReadOnlyList<byte[]>? records)
                    ? records
                    : Array.Empty<byte[]>();

            internal IReadOnlyList<byte[]> GetStyle9(
                SlideMasterPart masterPart) => _style9Records.TryGetValue(
                    masterPart.Uri.ToString(),
                    out IReadOnlyList<byte[]>? records)
                        ? records
                        : Array.Empty<byte[]>();
        }

        internal sealed class LegacyPptWriterFontCatalog {
            private readonly Dictionary<string, ushort> _indices =
                new(StringComparer.OrdinalIgnoreCase);
            private readonly List<KeyValuePair<ushort, string>> _added = new();
            private readonly int? _prototypeOffset;
            private int _nextIndex;

            internal LegacyPptWriterFontCatalog(LegacyPptRecord document) {
                LegacyPptRecord? collection = document.DescendantsAndSelf()
                    .FirstOrDefault(record =>
                        record.Type == RecordFontCollectionForWrite);
                _prototypeOffset = collection?.Offset;
                foreach (LegacyPptRecord atom in collection?.Children.Where(
                             child => child.Type == RecordFontEntityAtomForWrite)
                         ?? Enumerable.Empty<LegacyPptRecord>()) {
                    _nextIndex = Math.Max(_nextIndex, atom.Instance + 1);
                    if (atom.PayloadLength != 68) continue;
                    string typeface = atom.ReadUtf16Text(0, 64);
                    int terminator = typeface.IndexOf('\0');
                    if (terminator >= 0) typeface = typeface.Substring(0,
                        terminator);
                    if (typeface.Length == 0 || _indices.ContainsKey(typeface)) {
                        continue;
                    }
                    _indices.Add(typeface, atom.Instance);
                }
            }

            internal bool TryGetOrAdd(string? typeface, out ushort index,
                out string? reason) {
                index = 0;
                reason = null;
                string value = typeface ?? string.Empty;
                if (value.Length == 0 || string.IsNullOrWhiteSpace(value)) {
                    reason = "A text style contains an empty typeface name.";
                    return false;
                }
                if (!string.Equals(value, value.Trim(),
                        StringComparison.Ordinal)) {
                    reason = $"Typeface '{value}' contains leading or trailing whitespace that cannot be normalized losslessly.";
                    return false;
                }
                if (value.Length > 31 || value.IndexOf('\0') >= 0) {
                    reason = $"Typeface '{value}' does not fit the binary PowerPoint FontEntityAtom name field.";
                    return false;
                }
                if (_indices.TryGetValue(value, out index)) return true;
                if (_nextIndex >= ushort.MaxValue) {
                    reason = "The binary PowerPoint font index space is exhausted.";
                    return false;
                }
                index = checked((ushort)_nextIndex++);
                _indices.Add(value, index);
                _added.Add(new KeyValuePair<ushort, string>(index, value));
                return true;
            }

            internal bool HasAddedFonts => _added.Count > 0;
            internal bool HasPrototype => _prototypeOffset.HasValue;

            internal bool TryRewriteCollection(LegacyPptRecord record,
                out byte[] rewritten) {
                if (_prototypeOffset.HasValue
                    && record.Offset == _prototypeOffset.Value
                    && record.Type == RecordFontCollectionForWrite) {
                    rewritten = BuildCollection(record);
                    return true;
                }
                if (!_prototypeOffset.HasValue || record.Children.Count == 0) {
                    rewritten = record.CopyRecordBytes();
                    return false;
                }
                var children = new List<byte[]>(record.Children.Count);
                bool changed = false;
                foreach (LegacyPptRecord child in record.Children) {
                    if (!changed && TryRewriteCollection(child,
                            out byte[] rewrittenChild)) {
                        children.Add(rewrittenChild);
                        changed = true;
                    } else {
                        children.Add(child.CopyRecordBytes());
                    }
                }
                rewritten = changed
                    ? BuildRecord(record.Version, record.Instance, record.Type,
                        Concat(children))
                    : record.CopyRecordBytes();
                return changed;
            }

            internal byte[] BuildCollection() => BuildContainer(
                RecordFontCollectionForWrite, instance: 0,
                _added.Select(font => BuildFontEntityRecord(font.Key,
                    font.Value)));

            internal byte[] BuildCollection(LegacyPptRecord prototype) {
                var children = prototype.Children.Select(child =>
                    child.CopyRecordBytes()).ToList();
                children.AddRange(_added.Select(font =>
                    BuildFontEntityRecord(font.Key, font.Value)));
                return BuildContainer(RecordFontCollectionForWrite,
                    prototype.Instance, children);
            }

            private static byte[] BuildFontEntityRecord(ushort index,
                string typeface) {
                var payload = new byte[68];
                byte[] name = Encoding.Unicode.GetBytes(typeface + "\0");
                Buffer.BlockCopy(name, 0, payload, 0,
                    Math.Min(name.Length, 64));
                payload[64] = 0;
                payload[65] = 0;
                payload[66] = 0x04;
                payload[67] = 0x20;
                return BuildRecord(version: 0, index,
                    RecordFontEntityAtomForWrite, payload);
            }
        }
    }
}
