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
        private const ushort RecordFontEntityAtomForWrite = 0x0FB7;

        internal static bool CanWriteMasterTextStyles(
            PowerPointPresentation presentation, out string? reason) =>
            TryReadMasterTextStyles(presentation, Template.Value.Document,
                out _, out reason);

        private static bool TryReadMasterTextStyles(
            PowerPointPresentation presentation, LegacyPptRecord templateDocument,
            out LegacyPptWriterMasterTextStyleCatalog catalog,
            out string? reason) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            var fonts = new LegacyPptWriterFontCatalog(templateDocument);
            var records = new Dictionary<string, IReadOnlyList<byte[]>>(
                StringComparer.Ordinal);
            foreach (SlideMasterPart masterPart in presentation.OpenXmlDocument
                         .PresentationPart?.SlideMasterParts
                     ?? Enumerable.Empty<SlideMasterPart>()) {
                if (!TryBuildMasterTextStyleRecords(
                        masterPart.SlideMaster?.TextStyles, fonts,
                        out IReadOnlyList<byte[]> masterRecords, out reason)) {
                    catalog = new LegacyPptWriterMasterTextStyleCatalog(
                        records, fonts);
                    return false;
                }
                records.Add(masterPart.Uri.ToString(), masterRecords);
            }
            catalog = new LegacyPptWriterMasterTextStyleCatalog(records, fonts);
            reason = null;
            return true;
        }

        private static bool TryBuildMasterTextStyleRecords(P.TextStyles? styles,
            LegacyPptWriterFontCatalog fonts, out IReadOnlyList<byte[]> records,
            out string? reason) {
            var result = new List<byte[]>(3);
            records = result;
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
            return true;
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
            var levels = new Dictionary<int, A.TextParagraphPropertiesType>();
            foreach (OpenXmlElement child in style.ChildElements) {
                if (child is not A.TextParagraphPropertiesType properties
                    || !TryGetMasterTextStyleLevel(properties, out int level)
                    || levels.ContainsKey(level)) {
                    reason = $"The {style.LocalName} master text style contains unsupported or duplicate level content.";
                    return false;
                }
                levels.Add(level, properties);
            }
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
                        out reason)
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

            internal LegacyPptWriterMasterTextStyleCatalog(
                IReadOnlyDictionary<string, IReadOnlyList<byte[]>> records,
                LegacyPptWriterFontCatalog fonts) {
                _records = new ReadOnlyDictionary<string,
                    IReadOnlyList<byte[]>>(records.ToDictionary(pair => pair.Key,
                    pair => (IReadOnlyList<byte[]>)new ReadOnlyCollection<byte[]>(
                        pair.Value.Select(record => record.ToArray()).ToArray()),
                    StringComparer.Ordinal));
                Fonts = fonts;
            }

            internal LegacyPptWriterFontCatalog Fonts { get; }

            internal IReadOnlyList<byte[]> Get(SlideMasterPart masterPart) =>
                _records.TryGetValue(masterPart.Uri.ToString(),
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
