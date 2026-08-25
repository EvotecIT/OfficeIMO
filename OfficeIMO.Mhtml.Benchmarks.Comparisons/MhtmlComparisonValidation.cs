using System.Security.Cryptography;
using AngleSharp.Html.Parser;
using MimeKit;

namespace OfficeIMO.Mhtml.Benchmarks.Comparisons;

internal sealed record MhtmlResourceObservation(
    string ContentType,
    string ContentId,
    string ContentLocation,
    string FileName,
    int Length,
    string Sha256);

internal sealed record MhtmlObservation(
    string Subject,
    string RootContentId,
    string ContentLocation,
    string Html,
    int HtmlElementCount,
    IReadOnlyList<MhtmlResourceObservation> Resources);

internal sealed record MhtmlComparisonReport(
    string Scale,
    int OfficeIMOOutputBytes,
    int MimeKitOutputBytes,
    int ResourceCount,
    int DecodedResourceBytes);

internal static class MhtmlComparisonValidation {
    internal static IReadOnlyList<MhtmlComparisonReport> ValidateAll() =>
        MhtmlComparisonCorpus.ScaleNames.Select(Validate).ToArray();

    internal static MhtmlComparisonReport Validate(string scaleName) {
        MhtmlBenchmarkScale scale = MhtmlComparisonCorpus.Get(scaleName);
        MhtmlDocument officeModel = MhtmlComparisonCorpus.CreateOfficeDocument(scale);
        using MimeMessage mimeModel = MhtmlComparisonCorpus.CreateMimeMessage(scale);
        byte[] officeOutput = officeModel.ToBytes();
        byte[] mimeOutput = MhtmlComparisonCorpus.WriteMimeKit(mimeModel);

        MhtmlObservation expected = Expected(scale);
        RequireEqual(scaleName, "OfficeIMO output through OfficeIMO", expected, ObserveOffice(officeOutput));
        RequireEqual(scaleName, "OfficeIMO output through MimeKit", expected, ObserveMimeKit(officeOutput));
        RequireEqual(scaleName, "MimeKit output through OfficeIMO", expected, ObserveOffice(mimeOutput));
        RequireEqual(scaleName, "MimeKit output through MimeKit", expected, ObserveMimeKit(mimeOutput));

        return new MhtmlComparisonReport(
            scale.Name,
            officeOutput.Length,
            mimeOutput.Length,
            expected.Resources.Count,
            expected.Resources.Sum(resource => resource.Length));
    }

    internal static void ValidateOutput(MhtmlBenchmarkScale scale, byte[] output) {
        MhtmlObservation expected = Expected(scale);
        RequireEqual(scale.Name, "evidence output through OfficeIMO", expected, ObserveOffice(output));
        RequireEqual(scale.Name, "evidence output through MimeKit", expected, ObserveMimeKit(output));
    }

    internal static MhtmlDocument LoadOffice(byte[] input) {
        using var stream = new MemoryStream(input, writable: false);
        return MhtmlDocument.Load(stream);
    }

    internal static MhtmlMimeKitProjection LoadMimeKit(byte[] input) {
        var stream = new MemoryStream(input, writable: false);
        MimeMessage message;
        try {
            message = MimeMessage.Load(stream);
        } catch {
            stream.Dispose();
            throw;
        }

        try {
            MultipartRelated related = RequireRelated(message);
            TextPart root = RequireRoot(related);
            var parser = new HtmlParser();
            var htmlDocument = parser.ParseDocument(root.Text);
            byte[][] resources = related
                .Where(entity => !IsSelectedRoot(entity, root))
                .OfType<MimePart>()
                .Select(ReadPart)
                .ToArray();
            return new MhtmlMimeKitProjection(stream, message, htmlDocument, resources);
        } catch {
            message.Dispose();
            stream.Dispose();
            throw;
        }
    }

    private static MhtmlObservation Expected(MhtmlBenchmarkScale scale) {
        string html = MhtmlComparisonCorpus.CreateHtml(scale);
        return new MhtmlObservation(
            MhtmlComparisonCorpus.Subject(scale),
            MhtmlComparisonCorpus.RootContentId(scale),
            MhtmlComparisonCorpus.ContentLocation(scale),
            NormalizeHtml(html),
            CountElements(html),
            MhtmlComparisonCorpus.CreateResources(scale).Select(Observe).ToArray());
    }

    private static MhtmlObservation ObserveOffice(byte[] input) {
        MhtmlDocument document = LoadOffice(input);
        return new MhtmlObservation(
            document.Subject ?? string.Empty,
            document.RootContentId ?? string.Empty,
            document.ContentLocation ?? string.Empty,
            NormalizeHtml(document.Html),
            CountElements(document.Html),
            document.Resources.Select(resource => {
                byte[] content = resource.Content;
                return new MhtmlResourceObservation(
                    resource.ContentType,
                    resource.ContentId ?? string.Empty,
                    resource.ContentLocation ?? string.Empty,
                    resource.FileName ?? string.Empty,
                    content.Length,
                    Convert.ToHexString(SHA256.HashData(content)));
            }).ToArray());
    }

    private static MhtmlObservation ObserveMimeKit(byte[] input) {
        using var stream = new MemoryStream(input, writable: false);
        using MimeMessage message = MimeMessage.Load(stream);
        MultipartRelated related = RequireRelated(message);
        TextPart root = RequireRoot(related);
        return new MhtmlObservation(
            message.Subject ?? string.Empty,
            NormalizeContentId(root.ContentId),
            root.ContentLocation?.ToString() ?? HeaderValue(message, "Snapshot-Content-Location"),
            NormalizeHtml(root.Text),
            CountElements(root.Text),
            related.Where(entity => !IsSelectedRoot(entity, root)).OfType<MimePart>().Select(part => {
                byte[] content = ReadPart(part);
                return new MhtmlResourceObservation(
                    part.ContentType.MimeType,
                    NormalizeContentId(part.ContentId),
                    part.ContentLocation?.ToString() ?? string.Empty,
                    part.FileName ?? string.Empty,
                    content.Length,
                    Convert.ToHexString(SHA256.HashData(content)));
            }).ToArray());
    }

    private static MhtmlResourceObservation Observe(MhtmlResourceData resource) => new(
        resource.ContentType,
        resource.ContentId,
        resource.ContentLocation,
        resource.FileName,
        resource.Content.Length,
        Convert.ToHexString(SHA256.HashData(resource.Content)));

    private static MultipartRelated RequireRelated(MimeMessage message) =>
        message.Body as MultipartRelated
        ?? throw new InvalidDataException("MHTML root is not multipart/related.");

    private static TextPart RequireRoot(MultipartRelated related) =>
        related.Root as TextPart
        ?? throw new InvalidDataException("MHTML root is not an HTML text part.");

    private static bool IsSelectedRoot(MimeEntity entity, TextPart root) =>
        ReferenceEquals(entity, root)
        || entity is TextPart text
            && text.IsHtml
            && NormalizeContentId(text.ContentId) == NormalizeContentId(root.ContentId);

    private static byte[] ReadPart(MimePart part) {
        using var output = new MemoryStream();
        (part.Content ?? throw new InvalidDataException("MHTML resource content is missing.")).DecodeTo(output);
        return output.ToArray();
    }

    private static int CountElements(string html) {
        var parser = new HtmlParser();
        using var document = parser.ParseDocument(html);
        return document.All.Length;
    }

    private static string NormalizeHtml(string value) =>
        value.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n').TrimEnd();

    private static string NormalizeContentId(string? value) =>
        string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim().Trim('<', '>');

    private static string HeaderValue(MimeMessage message, string name) => message.Headers[name] ?? string.Empty;

    private static void RequireEqual(string scale, string path, MhtmlObservation expected, MhtmlObservation actual) {
        bool equal = expected.Subject == actual.Subject
            && expected.RootContentId == actual.RootContentId
            && expected.ContentLocation == actual.ContentLocation
            && expected.Html == actual.Html
            && expected.HtmlElementCount == actual.HtmlElementCount
            && expected.Resources.SequenceEqual(actual.Resources);
        if (!equal) {
            throw new InvalidOperationException(
                $"MHTML semantic validation failed for {scale}/{path}. " +
                $"Subject={expected.Subject == actual.Subject}, " +
                $"RootContentId={expected.RootContentId == actual.RootContentId} " +
                $"('{expected.RootContentId}'/'{actual.RootContentId}'), " +
                $"ContentLocation={expected.ContentLocation == actual.ContentLocation} " +
                $"('{expected.ContentLocation}'/'{actual.ContentLocation}'), " +
                $"Html={expected.Html == actual.Html} ({expected.Html.Length}/{actual.Html.Length}), " +
                $"Elements={expected.HtmlElementCount}/{actual.HtmlElementCount}, " +
                $"Resources={expected.Resources.Count}/{actual.Resources.Count}, " +
                $"ResourceEquality={expected.Resources.SequenceEqual(actual.Resources)}, " +
                $"ActualResourceIds=[{string.Join(",", actual.Resources.Select(resource => resource.ContentId))}].");
        }
    }
}

internal sealed class MhtmlMimeKitProjection : IDisposable {
    private readonly Stream _stream;

    internal MhtmlMimeKitProjection(Stream stream, MimeMessage message, object htmlDocument, IReadOnlyList<byte[]> resources) {
        _stream = stream;
        Message = message;
        HtmlDocument = htmlDocument;
        Resources = resources;
    }

    internal MimeMessage Message { get; }
    internal object HtmlDocument { get; }
    internal IReadOnlyList<byte[]> Resources { get; }

    public void Dispose() {
        Message.Dispose();
        if (HtmlDocument is IDisposable disposable) disposable.Dispose();
        _stream.Dispose();
    }
}
