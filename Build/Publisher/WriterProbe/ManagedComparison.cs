using System.Text;
using System.Text.Json;
using OfficeIMO.Publisher.Pdf;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Produces managed comparison artifacts without treating rendering success as native fidelity.</summary>
internal static class ManagedComparison {
    internal static object Render(string input, string directory) {
        string output = Path.GetFullPath(directory);
        if (Directory.Exists(output) || File.Exists(output)) throw new IOException("Choose a new comparison output directory.");
        string parent = Path.GetDirectoryName(output) ?? throw new ArgumentException("Output has no parent directory.");
        if (!Directory.Exists(parent)) throw new DirectoryNotFoundException(parent);
        using var source = new FileStream(input, FileMode.Open, FileAccess.Read, FileShare.Read);
        PublisherDocument publication = PublisherDocument.Load(source);
        source.Position = 0;
        string sourceHash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(source)).ToLowerInvariant();
        var conversion = publication.ToPdfDocumentResult();
        byte[] pdf = conversion.ToBytes();
        if (pdf.Length < 5 || Encoding.ASCII.GetString(pdf, 0, 5) != "%PDF-")
            throw new InvalidDataException("Managed PDF output has no PDF header.");
        Directory.CreateDirectory(output);
        var artifacts = new List<object>();
        Write("managed.pdf", pdf);
        for (int index = 0; index < publication.Pages.Count; index++) {
            var svg = publication.ToSvgResult(index);
            Write($"page-{index + 1:D3}.svg", Encoding.UTF8.GetBytes(svg.Value));
        }
        var normalized = new {
            schema = "officeimo.publisher.managed-comparison.v1",
            source = Path.GetFullPath(input), sourceSha256 = sourceHash,
            nativeAcceptance = "not-run", visualFidelity = "not-assessed",
            pages = publication.Pages.Select(Page).ToArray(),
            masters = publication.MasterPages.Select(Page).ToArray(),
            stories = publication.TextStories.Select(story => new { story.Id, story.Text, story.Paragraphs }).ToArray(),
            images = publication.Images.Select(image => new { image.Id, image.ContentType, image.ByteCount, sha256 = NativeInventory.Hash(image.GetBytes()) }).ToArray(),
            sourceReport = publication.ReadReport,
            conversionReports = conversion.ConversionReports.Select(report => new { report.HasLoss, report.FidelityDiagnostics }).ToArray(),
            artifacts
        };
        string json = JsonSerializer.Serialize(normalized, Program.JsonOptions);
        File.WriteAllText(Path.Combine(output, "managed.json"), json, new UTF8Encoding(false));
        return new { output, sourceSha256 = sourceHash, pages = publication.Pages.Count, stories = publication.TextStories.Count, nativeAcceptance = "not-run", visualFidelity = "not-assessed" };

        void Write(string name, byte[] bytes) {
            using var file = new FileStream(Path.Combine(output, name), FileMode.CreateNew, FileAccess.Write);
            file.Write(bytes);
            artifacts.Add(new { name, bytes = bytes.Length, sha256 = NativeInventory.Hash(bytes) });
        }
    }

    private static object Page(PublisherPage page) => new {
        page.Id, page.Name, page.MasterPageId, page.Width, page.Height,
        page.TextFrames, page.Tables,
        drawingElements = page.Drawing.Elements.Count
    };
}
