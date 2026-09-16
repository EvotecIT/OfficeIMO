using System.Text.Json;
using OfficeIMO.Html.Runtime;

if (args.Length < 2) throw new ArgumentException("Usage: OfficeIMO.Html.PublicFetchProbe <http(s)-url> <new-output-directory> [extra-allowed-host ...]");
Uri url = HtmlPublicResourceBroker.ValidateUrl(new Uri(args[0], UriKind.Absolute));
string outputDirectory = Path.GetFullPath(args[1]);
if (Directory.Exists(outputDirectory) || File.Exists(outputDirectory))
    throw new IOException("The probe output directory must not already exist.");
var broker = new HtmlPublicResourceBroker(new[] { url.IdnHost }.Concat(args.Skip(2)));
HtmlPublicResourceResult result = await broker.FetchAsync(url);
string html = HtmlPublicResourceBroker.DecodeUtf8Html(result.Resource);
Directory.CreateDirectory(outputDirectory);
await File.WriteAllBytesAsync(Path.Combine(outputDirectory, "response.html"), result.Resource.Content);
var manifest = new {
    requestedUrl = url.AbsoluteUri,
    finalUrl = result.Resource.FinalUrl.AbsoluteUri,
    result.FetchedAtUtc,
    connectedAddress = result.ConnectedAddress.ToString(),
    result.Resource.ContentType,
    result.Resource.StatusCode,
    byteCount = result.Resource.Length,
    result.Sha256,
    htmlCharacters = html.Length,
    redirects = result.Redirects.Select(hop => new {
        from = hop.From.AbsoluteUri, to = hop.To.AbsoluteUri,
        hop.StatusCode, connectedAddress = hop.ConnectedAddress.ToString()
    }).ToArray()
};
await File.WriteAllTextAsync(Path.Combine(outputDirectory, "manifest.json"),
    JsonSerializer.Serialize(manifest, new JsonSerializerOptions { WriteIndented = true }));
Console.WriteLine(result.Resource.FinalUrl + " " + result.Resource.Length + " bytes SHA-256 " + result.Sha256);
