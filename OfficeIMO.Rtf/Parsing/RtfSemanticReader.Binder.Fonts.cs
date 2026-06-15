using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

internal static partial class RtfSemanticReader {
    private sealed partial class Binder {
        private static IReadOnlyList<RtfFont> ReadFontTable(RtfGroup root, int ansiCodePage, int unicodeSkipCount) {
            RtfGroup? table = root.Children.OfType<RtfGroup>().FirstOrDefault(group => group.Destination == "fonttbl");
            if (table == null) return new[] { new RtfFont(0, "Calibri") };

            var fonts = new List<RtfFont>();
            foreach (RtfGroup fontGroup in table.Children.OfType<RtfGroup>()) {
                FontTableEntry entry = ReadFontTableEntry(fontGroup, ansiCodePage, unicodeSkipCount);
                if (entry.Id.HasValue && !string.IsNullOrWhiteSpace(entry.Name)) {
                    fonts.Add(new RtfFont(entry.Id.Value, entry.Name) {
                        Family = entry.Family,
                        Charset = entry.Charset,
                        Pitch = entry.Pitch,
                        CodePage = entry.CodePage,
                        Bias = entry.Bias,
                        AlternateName = EmptyToNull(entry.AlternateName ?? string.Empty),
                        Panose = EmptyToNull(entry.Panose ?? string.Empty),
                        NonTaggedName = EmptyToNull(entry.NonTaggedName ?? string.Empty),
                        Embedding = entry.Embedding
                    });
                }
            }

            return fonts.Count == 0 ? new[] { new RtfFont(0, "Calibri") } : fonts;
        }

        private static FontTableEntry ReadFontTableEntry(RtfGroup fontGroup, int ansiCodePage, int unicodeSkipCount) {
            var entry = new FontTableEntry();

            foreach (RtfNode node in fontGroup.Children) {
                if (node is RtfControlWord control) {
                    ApplyFontTableControl(control, entry);
                } else if (node is RtfGroup childGroup) {
                    ApplyFontTableDestination(childGroup, entry, ansiCodePage, unicodeSkipCount);
                }
            }

            entry.Name = CollectDirectPlainText(fontGroup.Children, ansiCodePage, unicodeSkipCount).Trim().TrimEnd(';').Trim();
            return entry;
        }

        private static void ApplyFontTableControl(RtfControlWord control, FontTableEntry entry) {
            switch (control.Name) {
                case "f":
                    if (control.Parameter.HasValue) {
                        entry.Id = control.Parameter.Value;
                    }
                    break;
                case "fnil":
                    entry.Family = RtfFontFamily.Nil;
                    break;
                case "froman":
                    entry.Family = RtfFontFamily.Roman;
                    break;
                case "fswiss":
                    entry.Family = RtfFontFamily.Swiss;
                    break;
                case "fmodern":
                    entry.Family = RtfFontFamily.Modern;
                    break;
                case "fscript":
                    entry.Family = RtfFontFamily.Script;
                    break;
                case "fdecor":
                    entry.Family = RtfFontFamily.Decorative;
                    break;
                case "ftech":
                    entry.Family = RtfFontFamily.Technical;
                    break;
                case "fbidi":
                    entry.Family = RtfFontFamily.Bidirectional;
                    break;
                case "fcharset":
                    entry.Charset = control.Parameter;
                    break;
                case "fprq":
                    entry.Pitch = control.Parameter;
                    break;
                case "cpg":
                    entry.CodePage = control.Parameter;
                    break;
                case "fbias":
                    entry.Bias = control.Parameter;
                    break;
            }
        }

        private static void ApplyFontTableDestination(RtfGroup group, FontTableEntry entry, int ansiCodePage, int unicodeSkipCount) {
            string value = CollectPlainText(group, ansiCodePage, unicodeSkipCount).Trim();
            switch (group.Destination) {
                case "falt":
                    entry.AlternateName = value;
                    break;
                case "panose":
                    entry.Panose = value;
                    break;
                case "fname":
                    entry.NonTaggedName = value;
                    break;
                case "fontemb":
                    entry.Embedding = ReadFontEmbedding(group, ansiCodePage, unicodeSkipCount);
                    break;
            }
        }

        private static RtfFontEmbedding ReadFontEmbedding(RtfGroup group, int ansiCodePage, int unicodeSkipCount) {
            var embedding = new RtfFontEmbedding();
            var data = new List<byte>();

            foreach (RtfNode node in group.Children) {
                if (node is RtfControlWord control) {
                    ApplyFontEmbeddingControl(control, embedding);
                } else if (node is RtfGroup childGroup) {
                    if (childGroup.Destination == "fontfile") {
                        ReadFontFile(childGroup, embedding, ansiCodePage, unicodeSkipCount);
                    } else {
                        ReadFontEmbeddingData(childGroup, data);
                    }
                } else if (node is RtfBinary binary) {
                    data.AddRange(binary.Data);
                } else if (node is RtfText text) {
                    AppendHexBytes(text.Text, data);
                }
            }

            embedding.Data = data.ToArray();
            return embedding;
        }

        private static void ApplyFontEmbeddingControl(RtfControlWord control, RtfFontEmbedding embedding) {
            switch (control.Name) {
                case "ftnil":
                    embedding.Type = RtfEmbeddedFontType.Unknown;
                    break;
                case "fttruetype":
                    embedding.Type = RtfEmbeddedFontType.TrueType;
                    break;
            }
        }

        private static void ReadFontFile(RtfGroup group, RtfFontEmbedding embedding, int ansiCodePage, int unicodeSkipCount) {
            foreach (RtfControlWord control in group.Children.OfType<RtfControlWord>()) {
                if (control.Name == "cpg") {
                    embedding.FileCodePage = control.Parameter;
                }
            }

            embedding.FileName = EmptyToNull(CollectPlainText(group, ansiCodePage, unicodeSkipCount).Trim());
        }

        private static void ReadFontEmbeddingData(RtfGroup group, List<byte> data) {
            foreach (RtfNode node in group.Children) {
                if (node is RtfBinary binary) {
                    data.AddRange(binary.Data);
                } else if (node is RtfText text) {
                    AppendHexBytes(text.Text, data);
                } else if (node is RtfGroup childGroup) {
                    ReadFontEmbeddingData(childGroup, data);
                }
            }
        }

        private sealed class FontTableEntry {
            public int? Id { get; set; }
            public string Name { get; set; } = string.Empty;
            public RtfFontFamily? Family { get; set; }
            public int? Charset { get; set; }
            public int? Pitch { get; set; }
            public int? CodePage { get; set; }
            public int? Bias { get; set; }
            public string? AlternateName { get; set; }
            public string? Panose { get; set; }
            public string? NonTaggedName { get; set; }
            public RtfFontEmbedding? Embedding { get; set; }
        }
    }
}
