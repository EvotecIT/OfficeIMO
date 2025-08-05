using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using OfficeIMO.Word;

namespace OfficeIMO.Markdown {
    /// <summary>
    /// Converts Word documents into Markdown text without relying on HTML or external tools.
    /// </summary>
    public partial class WordToMarkdownConverter : IWordConverter {
        /// <summary>
        /// Converts a DOCX document from the provided stream into Markdown text.
        /// </summary>
        /// <param name="input">Stream containing DOCX content.</param>
        /// <param name="options">Conversion options.</param>
        /// <returns>Markdown representation of the document.</returns>
        public static string Convert(Stream input, WordToMarkdownOptions? options = null) {
            if (input == null) {
                throw new ConversionException($"{nameof(input)} cannot be null.");
            }

            options ??= new WordToMarkdownOptions();
            StringBuilder sb = new StringBuilder();

            using var word = WordprocessingDocument.Open(input, false);
            var body = word.MainDocumentPart?.Document.Body;
            if (body == null) return string.Empty;

            foreach (var paragraph in body.Elements<Paragraph>()) {
                string line = GetParagraphText(paragraph);
                if (string.IsNullOrEmpty(line)) {
                    continue;
                }

                var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                if (styleId != null && styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase)) {
                    if (int.TryParse(styleId.Substring("Heading".Length), out int level) && level > 0 && level <= 6) {
                        sb.Append('#', level).Append(' ').Append(line).AppendLine();
                        sb.AppendLine();
                        continue;
                    }
                }

                if (paragraph.ParagraphProperties?.NumberingProperties != null) {
                    bool bullet = ListParser.IsBullet(paragraph, word.MainDocumentPart!);
                    sb.Append(bullet ? "- " : "1. ").Append(line).AppendLine();
                    continue;
                }

                sb.AppendLine(line);
                sb.AppendLine();
            }

            return sb.ToString().TrimEnd();
        }

        public void Convert(Stream input, Stream output, IConversionOptions options) {
            string markdown = Convert(input, options as WordToMarkdownOptions);
            using StreamWriter writer = new StreamWriter(
                output,
                Encoding.UTF8,
                bufferSize: 1024,
                leaveOpen: true);
            writer.Write(markdown);
            writer.Flush();
        }

        public async Task ConvertAsync(Stream input, Stream output, IConversionOptions options, CancellationToken cancellationToken = default) {
            string markdown = Convert(input, options as WordToMarkdownOptions);
            using StreamWriter writer = new StreamWriter(
                output,
                Encoding.UTF8,
                bufferSize: 1024,
                leaveOpen: true);
#if NET8_0_OR_GREATER
            await writer.WriteAsync(markdown.AsMemory(), cancellationToken).ConfigureAwait(false);
            await writer.FlushAsync(cancellationToken).ConfigureAwait(false);
#else
            await writer.WriteAsync(markdown).ConfigureAwait(false);
            await writer.FlushAsync().ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
#endif
        }
    }
}
