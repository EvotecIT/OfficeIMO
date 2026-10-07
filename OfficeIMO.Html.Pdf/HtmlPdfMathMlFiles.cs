using System.Collections.Generic;
using System.Threading;
using System.Security.Cryptography;
using System.Text;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

// The adapter selects MathML payload and metadata; PDF owns filespecs, associations,
// deduplication, version requirements and conformance/resource gates.
internal sealed class HtmlPdfMathMlFiles {
    private readonly Dictionary<string, PdfCore.PdfEmbeddedFile> _files = new(StringComparer.Ordinal);
    private readonly DateTimeOffset? _modificationDate;
    internal HtmlPdfMathMlFiles(DateTimeOffset? modificationDate) => _modificationDate = modificationDate;

    internal void Attach(PdfCore.PdfCanvasStructureOptions options, HtmlMathMlSource source, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!_files.TryGetValue(source.MathMl, out PdfCore.PdfEmbeddedFile? file)) {
            byte[] bytes = Encoding.UTF8.GetBytes(source.MathMl);
            using var sha = SHA256.Create();
            string hash = BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", string.Empty).ToLowerInvariant();
            cancellationToken.ThrowIfCancellationRequested();
            file = new PdfCore.PdfEmbeddedFile("mathml-" + hash + ".mathml", bytes, "application/mathml+xml",
                PdfCore.PdfAssociatedFileRelationship.Supplement, "MathML source of the associated formula", modificationDate: _modificationDate);
            _files.Add(source.MathMl, file);
        }
        options.AddAssociatedFile(file);
    }
}
