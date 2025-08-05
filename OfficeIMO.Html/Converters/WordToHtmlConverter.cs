using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Html {
    /// <summary>
    /// Converts WordprocessingDocument content into simple HTML fragments.
    /// </summary>
    public partial class WordToHtmlConverter : IWordConverter {
        /// <summary>
        /// Converts a DOCX contained in the provided stream into HTML.
        /// </summary>
        /// <param name="docxStream">Stream containing DOCX content.</param>
        /// <param name="options">Conversion options.</param>
        /// <returns>Generated HTML string.</returns>
        public static string Convert(Stream docxStream, WordToHtmlOptions? options = null) {
            if (docxStream == null) {
                throw new ConversionException($"{nameof(docxStream)} cannot be null.");
            }

            options ??= new WordToHtmlOptions();

            using WordprocessingDocument document = WordprocessingDocument.Open(docxStream, false);
            StringBuilder sb = new StringBuilder();
            sb.Append("<html><body>");

            Dictionary<int, bool> listTypes = ListParser.GetListTypes(document.MainDocumentPart!);
            AppendElements(document.MainDocumentPart!.Document.Body!.ChildElements, sb, options, listTypes, document.MainDocumentPart);

            sb.Append("</body></html>");
            return sb.ToString();
        }

        public void Convert(Stream input, Stream output, IConversionOptions options) {
            string html = Convert(input, options as WordToHtmlOptions);
            using StreamWriter writer = new StreamWriter(
                output,
                Encoding.UTF8,
                bufferSize: 1024,
                leaveOpen: true);
            writer.Write(html);
            writer.Flush();
        }

        public async Task ConvertAsync(Stream input, Stream output, IConversionOptions options, CancellationToken cancellationToken = default) {
            string html = Convert(input, options as WordToHtmlOptions);
            using StreamWriter writer = new StreamWriter(
                output,
                Encoding.UTF8,
                bufferSize: 1024,
                leaveOpen: true);
#if NET8_0_OR_GREATER
            await writer.WriteAsync(html.AsMemory(), cancellationToken).ConfigureAwait(false);
            await writer.FlushAsync(cancellationToken).ConfigureAwait(false);
#else
            await writer.WriteAsync(html).ConfigureAwait(false);
            await writer.FlushAsync().ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
#endif
        }
    }
}