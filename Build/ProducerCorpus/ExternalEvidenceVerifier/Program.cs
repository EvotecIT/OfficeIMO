using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml;
using OfficeIMO.Html;
using OfficeIMO.OpenDocument;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Markdown;
using OfficeIMO.Word.Rtf;

if (args.Length != 2 || (args[0] != "rtf" && args[0] != "odf")) {
    Console.Error.WriteLine("Usage: ExternalEvidenceVerifier <rtf|odf> <manifest-path>");
    return 2;
}

string mode = args[0];
string manifestPath = Path.GetFullPath(args[1]);
Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
using JsonDocument manifest = JsonDocument.Parse(File.ReadAllBytes(manifestPath));
using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(90) };

try {
    object[] evidence = mode == "rtf"
        ? await VerifyRtfAsync(manifest.RootElement, client)
        : await VerifyOdfAsync(manifest.RootElement, client);
    Console.WriteLine(JsonSerializer.Serialize(evidence, new JsonSerializerOptions { WriteIndented = true }));
    return 0;
} catch (Exception exception) {
    Console.Error.WriteLine(exception.ToString());
    return 1;
}

static async Task<object[]> VerifyRtfAsync(JsonElement manifest, HttpClient client) {
    var output = new List<object>();
    foreach (JsonElement artifact in manifest.GetProperty("externalArtifacts").EnumerateArray()) {
        string id = artifact.GetProperty("id").GetString()!;
        string sourceUrl = artifact.GetProperty("sourceUrl").GetString()!;
        byte[] bytes = await client.GetByteArrayAsync(sourceUrl).ConfigureAwait(false);
        Require(bytes.LongLength == artifact.GetProperty("bytes").GetInt64(), id + " byte length changed");
        Require(Hash(bytes) == artifact.GetProperty("sha256").GetString(), id + " SHA-256 changed");

        string header = Encoding.GetEncoding(1252).GetString(bytes, 0, Math.Min(1024, bytes.Length));
        foreach (JsonElement fragment in artifact.GetProperty("requiredHeaderFragments").EnumerateArray()) {
            Require(header.Contains(fragment.GetString()!, StringComparison.Ordinal), id + " header evidence changed");
        }

        RtfReadResult read = RtfDocument.Load(bytes, RtfReadOptions.CreateUntrustedProfile());
        RtfToHtmlResult html = read.Document.ToHtmlResult(RtfToHtmlOptions.CreateWebSafeProfile());
        RtfConversionResult<string> markdown = read.Document.ToMarkdownResult();
        RtfConversionResult<OfficeIMO.Word.WordDocument> word = read.ToWordDocumentResult(sourceUrl);
        using (word.Value) {
            Require(!string.IsNullOrWhiteSpace(html.Value), id + " produced empty safe HTML");
            Require(!string.IsNullOrWhiteSpace(markdown.Value), id + " produced empty Markdown");
            Require(word.Succeeded, id + " failed the Word bridge");
            output.Add(new {
                id,
                producer = artifact.GetProperty("producer").GetString(),
                bytes = bytes.Length,
                sha256 = Hash(bytes),
                readDiagnostics = read.Diagnostics.Count,
                safeHtmlCharacters = html.Value.Length,
                markdownCharacters = markdown.Value.Length,
                wordBridgeDiagnostics = word.Report.Diagnostics.Count,
                productPaths = new[] { "RtfDocument.Load(untrusted)", "ToHtmlResult(web-safe)", "ToMarkdownResult", "RtfReadResult.ToWordDocumentResult" }
            });
        }
    }
    return output.ToArray();
}

static async Task<object[]> VerifyOdfAsync(JsonElement manifest, HttpClient client) {
    var output = new List<object>();
    foreach (JsonElement artifact in manifest.GetProperty("externalArtifacts").EnumerateArray()) {
        string id = artifact.GetProperty("id").GetString()!;
        byte[] bytes = await client.GetByteArrayAsync(artifact.GetProperty("sourceUrl").GetString()!).ConfigureAwait(false);
        int minBytes = artifact.GetProperty("minBytes").GetInt32();
        int maxBytes = artifact.GetProperty("maxBytes").GetInt32();
        Require(bytes.Length >= minBytes && bytes.Length <= maxBytes, id + " package size left its recorded range");

        var loadOptions = new OdfLoadOptions {
            MaxPackageBytes = Math.Max(maxBytes, 8 * 1024 * 1024),
            MaxEntries = 2048,
            MaxEntryUncompressedBytes = 32 * 1024 * 1024,
            MaxTotalUncompressedBytes = 96 * 1024 * 1024,
            MaxCompressionRatio = 200,
            MaxDepth = 24,
            MaxXmlCharacters = 32 * 1024 * 1024,
            MaxXmlDepth = 192
        };
        using var input = new MemoryStream(bytes, writable: false);
        OdfDocument document = OdfDocument.Load(input, loadOptions);
        Require(document.Validate().IsValid, id + " failed OfficeIMO validation");
        using var savedStream = new MemoryStream();
        OdfSaveResult saved = document.Save(savedStream);
        byte[] savedBytes = saved.RequireNoLoss();
        using var reopenedStream = new MemoryStream(savedBytes, writable: false);
        OdfDocument reopened = OdfDocument.Load(reopenedStream, loadOptions);
        Require(reopened.Validate().IsValid, id + " failed OfficeIMO save/reopen validation");

        (int paragraphCount, string semanticHash) = ReadOdfSemanticEvidence(bytes);
        Require(paragraphCount == artifact.GetProperty("paragraphCount").GetInt32(), id + " paragraph count changed");
        Require(semanticHash == artifact.GetProperty("semanticTextSha256").GetString(), id + " semantic hash changed to " + semanticHash);
        output.Add(new {
            id,
            producer = artifact.GetProperty("producer").GetString(),
            bytes = bytes.Length,
            paragraphCount,
            semanticTextSha256 = semanticHash,
            savedBytes = savedBytes.Length,
            savedWithoutLoss = !saved.HasLoss,
            productPaths = new[] { "OdfDocument.Load(bounded)", "Validate", "Save(RequireNoLoss)", "Load(saved)" }
        });
    }
    return output.ToArray();
}

static (int ParagraphCount, string Hash) ReadOdfSemanticEvidence(byte[] bytes) {
    using var stream = new MemoryStream(bytes, writable: false);
    using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
    ZipArchiveEntry content = archive.GetEntry("content.xml")
        ?? throw new InvalidDataException("OpenDocument evidence has no content.xml");
    var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null };
    using Stream contentStream = content.Open();
    using XmlReader reader = XmlReader.Create(contentStream, settings);
    var document = new XmlDocument { XmlResolver = null };
    document.Load(reader);
    var namespaces = new XmlNamespaceManager(document.NameTable);
    namespaces.AddNamespace("text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0");
    XmlNodeList nodes = document.SelectNodes("//text:p|//text:h", namespaces)
        ?? throw new InvalidDataException("OpenDocument evidence has no text paragraphs");
    string[] paragraphs = nodes.Cast<XmlNode>().Select(node => node.InnerText).ToArray();
    return (paragraphs.Length, Hash(Encoding.UTF8.GetBytes(string.Join("\n", paragraphs))));
}

static string Hash(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

static void Require(bool condition, string message) {
    if (!condition) throw new InvalidDataException(message);
}
