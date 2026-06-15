using System.Globalization;

namespace OfficeIMO.Rtf.Html;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private void ApplyFontTable(Dictionary<string, string> values) {
            var fonts = new List<RtfFont>();
            for (int index = 0; ; index++) {
                string prefix = "font." + index.ToString(CultureInfo.InvariantCulture);
                int? id = ReadInt(values, prefix + ".id");
                string? name = ReadString(values, prefix + ".name");
                if (!id.HasValue || string.IsNullOrWhiteSpace(name)) {
                    break;
                }

                var font = new RtfFont(id.Value, name!) {
                    Family = ReadEnum<RtfFontFamily>(values, prefix + ".family"),
                    Charset = ReadInt(values, prefix + ".charset"),
                    Pitch = ReadInt(values, prefix + ".pitch"),
                    CodePage = ReadInt(values, prefix + ".codePage"),
                    Bias = ReadInt(values, prefix + ".bias"),
                    AlternateName = ReadString(values, prefix + ".alternateName"),
                    Panose = ReadString(values, prefix + ".panose"),
                    NonTaggedName = ReadString(values, prefix + ".nonTaggedName"),
                    Embedding = ReadFontEmbedding(values, prefix + ".embedding")
                };
                fonts.Add(font);
            }

            if (fonts.Count > 0) {
                _document.ReplaceFonts(fonts);
            }
        }

        private static RtfFontEmbedding? ReadFontEmbedding(Dictionary<string, string> values, string prefix) {
            RtfEmbeddedFontType? type = ReadEnum<RtfEmbeddedFontType>(values, prefix + ".type");
            string? fileName = ReadString(values, prefix + ".fileName");
            int? fileCodePage = ReadInt(values, prefix + ".fileCodePage");
            byte[] data = DecodeBytes(ReadString(values, prefix + ".data"));
            if (!type.HasValue && string.IsNullOrWhiteSpace(fileName) && !fileCodePage.HasValue && data.Length == 0) {
                return null;
            }

            return new RtfFontEmbedding {
                Type = type ?? RtfEmbeddedFontType.Unknown,
                FileName = fileName,
                FileCodePage = fileCodePage,
                Data = data
            };
        }

        private static string? ReadString(Dictionary<string, string> values, string key) {
            return values.TryGetValue(key, out string? value) && value.Length > 0
                ? value
                : null;
        }
    }
}
