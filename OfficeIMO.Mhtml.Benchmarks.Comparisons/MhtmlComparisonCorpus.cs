using System.Text;
using MimeKit;

namespace OfficeIMO.Mhtml.Benchmarks.Comparisons;

internal sealed record MhtmlBenchmarkScale(
    string Name,
    int HtmlCharacters,
    int ResourceCount,
    int ResourceBytes);

internal static class MhtmlComparisonCorpus {
    internal static IReadOnlyList<string> ScaleNames { get; } = ["Small", "Normal", "Large"];

    internal static MhtmlBenchmarkScale Get(string name) => name switch {
        "Small" => new MhtmlBenchmarkScale("Small", 4 * 1024, 2, 4 * 1024),
        "Normal" => new MhtmlBenchmarkScale("Normal", 64 * 1024, 16, 32 * 1024),
        "Large" => new MhtmlBenchmarkScale("Large", 512 * 1024, 64, 128 * 1024),
        _ => throw new ArgumentException("Unknown MHTML benchmark scale: " + name, nameof(name))
    };

    internal static MhtmlDocument CreateOfficeDocument(MhtmlBenchmarkScale scale) {
        MhtmlResource[] resources = CreateResources(scale)
            .Select(resource => new MhtmlResource(
                resource.Content,
                resource.ContentType,
                resource.ContentId,
                resource.ContentLocation,
                resource.FileName))
            .ToArray();
        return new MhtmlDocument(
            CreateHtml(scale),
            resources,
            ContentLocation(scale),
            RootContentId(scale),
            Subject(scale));
    }

    internal static MimeMessage CreateMimeMessage(MhtmlBenchmarkScale scale) {
        var message = new MimeMessage { Subject = Subject(scale) };
        var related = new MultipartRelated();
        var root = new TextPart("html") {
            Text = CreateHtml(scale),
            ContentId = RootContentId(scale),
            ContentLocation = new Uri(ContentLocation(scale))
        };
        related.Add(root);
        foreach (MhtmlResourceData resource in CreateResources(scale)) {
            string[] mediaType = resource.ContentType.Split('/');
            var part = new MimePart(mediaType[0], mediaType[1]) {
                Content = new MimeContent(new MemoryStream(resource.Content, writable: false)),
                ContentTransferEncoding = ContentEncoding.Base64,
                ContentDisposition = new ContentDisposition(ContentDisposition.Inline),
                ContentId = resource.ContentId,
                ContentLocation = new Uri(resource.ContentLocation, UriKind.Relative),
                FileName = resource.FileName
            };
            related.Add(part);
        }
        related.Root = root;
        message.Body = related;
        message.Headers.Add("Snapshot-Content-Location", ContentLocation(scale));
        return message;
    }

    internal static byte[] WriteMimeKit(MimeMessage message) {
        using var output = new MemoryStream();
        message.WriteTo(output);
        return output.ToArray();
    }

    internal static string Subject(MhtmlBenchmarkScale scale) => "OfficeIMO MHTML comparison " + scale.Name;
    internal static string RootContentId(MhtmlBenchmarkScale scale) => "root-" + scale.Name.ToLowerInvariant() + "@example.test";
    internal static string ContentLocation(MhtmlBenchmarkScale scale) =>
        "https://archive.example.test/" + scale.Name.ToLowerInvariant() + "/index.html";

    internal static string CreateHtml(MhtmlBenchmarkScale scale) {
        var builder = new StringBuilder(scale.HtmlCharacters + 1024);
        builder.Append("<!doctype html><html><head><title>").Append(Subject(scale))
            .Append("</title></head><body><main>");
        int paragraph = 0;
        while (builder.Length < scale.HtmlCharacters) {
            int resource = paragraph % scale.ResourceCount;
            builder.Append("<section><h2>Section ").Append(paragraph)
                .Append("</h2><p>Deterministic MHTML content with Unicode zażółć and semantic text ")
                .Append(paragraph).Append(".</p><img alt=\"resource ").Append(resource)
                .Append("\" src=\"cid:resource-").Append(resource).Append("@example.test\"></section>");
            paragraph++;
        }
        builder.Append("</main></body></html>");
        return builder.ToString();
    }

    internal static IReadOnlyList<MhtmlResourceData> CreateResources(MhtmlBenchmarkScale scale) {
        var resources = new MhtmlResourceData[scale.ResourceCount];
        for (int index = 0; index < resources.Length; index++) {
            byte[] content = new byte[scale.ResourceBytes];
            for (int offset = 0; offset < content.Length; offset++) {
                content[offset] = (byte)((offset * 31 + index * 17 + 11) & 0xFF);
            }
            resources[index] = new MhtmlResourceData(
                content,
                "image/png",
                $"resource-{index}@example.test",
                $"assets/resource-{index:D3}.png",
                $"resource-{index:D3}.png");
        }
        return resources;
    }
}

internal sealed record MhtmlResourceData(
    byte[] Content,
    string ContentType,
    string ContentId,
    string ContentLocation,
    string FileName);
