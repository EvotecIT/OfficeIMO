using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Word;

namespace OfficeIMO.Word.Pdf {
    /// <summary>
    /// Provides stream based conversion from Word documents to PDF.
    /// </summary>
    public class WordPdfConverter : IWordConverter {
        public void Convert(Stream input, Stream output, IConversionOptions options) {
            using WordDocument document = WordDocument.Load(input);
            document.SaveAsPdf(output, options as PdfSaveOptions);
        }

        public async Task ConvertAsync(Stream input, Stream output, IConversionOptions options, CancellationToken cancellationToken = default) {
            using WordDocument document = WordDocument.Load(input);
            document.SaveAsPdf(output, options as PdfSaveOptions);
            await output.FlushAsync(cancellationToken).ConfigureAwait(false);
        }
    }
}
