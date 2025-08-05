using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Html {
    /// <summary>
    /// Converts simple HTML fragments into WordprocessingDocument instances.
    /// </summary>
    public partial class HtmlToWordConverter : IWordConverter {
        private static readonly Regex _urlRegex = new("((?:https?|ftp)://[^\\s]+)", RegexOptions.IgnoreCase);
        /// <summary>
        /// Converts provided HTML string into a DOCX document written to the specified stream.
        /// </summary>
        /// <param name="html">HTML content to convert. It should be a valid XHTML fragment.</param>
        /// <param name="output">Stream where DOCX content will be written.</param>
        /// <param name="options">Conversion options.</param>
        /// <param name="cancellationToken">Token used to cancel the operation.</param>
        public static void Convert(string html, Stream output, HtmlToWordOptions? options = null, CancellationToken cancellationToken = default) {
            if (html == null) {
                throw new ConversionException($"{nameof(html)} cannot be null.");
            }
            if (output == null) {
                throw new ConversionException($"{nameof(output)} cannot be null.");
            }

            options ??= new HtmlToWordOptions();

            using var document = WordDocument.Create();
            options.ApplyDefaults(document);
            WordprocessingDocument wordDoc = document._wordprocessingDocument;
            MainDocumentPart mainPart = wordDoc.MainDocumentPart;
            Body body = mainPart.Document.Body;

            // add numbering definitions for ordered and unordered lists using shared Word logic
            NumberingDefinitionsPart numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
            numberingPart.Numbering = new Numbering();
            Numbering numbering = WordListStyles.CreateDefaultNumberingDefinitions(wordDoc, out int bulletNumberId, out int orderedNumberId);
            numberingPart.Numbering = numbering;

            XDocument xdoc = XDocument.Parse("<root>" + html + "</root>");

            foreach (XElement element in xdoc.Root!.Elements()) {
                cancellationToken.ThrowIfCancellationRequested();
                AppendBlockElement(body, element, options, 0, bulletNumberId, orderedNumberId, mainPart, cancellationToken);
            }

            mainPart.Document.Save();
            document.Save(output);
        }


        public void Convert(Stream input, Stream output, IConversionOptions options) {
            if (input == null) {
                throw new ConversionException($"{nameof(input)} cannot be null.");
            }
            using StreamReader reader = new StreamReader(
                input,
                Encoding.UTF8,
                detectEncodingFromByteOrderMarks: true,
                bufferSize: 1024,
                leaveOpen: true);
            string html = reader.ReadToEnd();
            Convert(html, output, options as HtmlToWordOptions);
        }

        public async Task ConvertAsync(Stream input, Stream output, IConversionOptions options, CancellationToken cancellationToken = default) {
            if (input == null) {
                throw new ConversionException($"{nameof(input)} cannot be null.");
            }
            using StreamReader reader = new StreamReader(
                input,
                Encoding.UTF8,
                detectEncodingFromByteOrderMarks: true,
                bufferSize: 1024,
                leaveOpen: true);
            string html;
#if NET8_0_OR_GREATER
            html = await reader.ReadToEndAsync(cancellationToken).ConfigureAwait(false);
#else
            html = await reader.ReadToEndAsync().ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
#endif
            Convert(html, output, options as HtmlToWordOptions, cancellationToken);
            await output.FlushAsync(cancellationToken).ConfigureAwait(false);
        }
    }
}