using System.Text;
using Microsoft.Playwright;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>
/// Keeps font bytes already loaded by the browser in a CDP MHTML snapshot that omitted them.
/// This is limited to the opt-in evidence capture path; it never fetches another URL.
/// </summary>
internal static class HtmlMhtmlFontCapture {
    private const int MaximumFonts = 16;
    private const int MaximumFontBytes = 2 * 1024 * 1024;
    private const int MaximumTotalBytes = 8 * 1024 * 1024;

    internal sealed record Result(string ArchiveText, int Observed, int Added, int AddedBytes, int AlreadyArchived);

    internal static async Task<Result> AppendLoadedFontsAsync(string snapshot, IReadOnlyCollection<IResponse> responses) {
        ArgumentNullException.ThrowIfNull(snapshot);
        ArgumentNullException.ThrowIfNull(responses);
        IResponse[] fonts = responses.GroupBy(response => response.Url, StringComparer.Ordinal)
            .Select(group => group.Last()).OrderBy(response => response.Url, StringComparer.Ordinal).ToArray();
        if (fonts.Length > MaximumFonts)
            throw new InvalidDataException($"Live page loaded {fonts.Length} font responses; the capture limit is {MaximumFonts}.");
        if (fonts.Length == 0) return new Result(snapshot, 0, 0, 0, 0);

        int closingStart = snapshot.LastIndexOf("\r\n--", StringComparison.Ordinal);
        if (closingStart < 0) throw new InvalidDataException("MHTML snapshot has no closing multipart boundary.");
        int closingEnd = snapshot.IndexOf("\r\n", closingStart + 2, StringComparison.Ordinal);
        if (closingEnd < 0 || closingEnd != snapshot.Length - 2)
            throw new InvalidDataException("MHTML snapshot closing boundary is not the final line.");
        string closing = snapshot[(closingStart + 2)..closingEnd];
        if (!closing.StartsWith("--", StringComparison.Ordinal) || !closing.EndsWith("--", StringComparison.Ordinal))
            throw new InvalidDataException("MHTML snapshot has an invalid closing boundary.");
        string boundary = closing[..^2];
        if (snapshot.IndexOf("\r\n" + boundary + "\r\n", StringComparison.Ordinal) < 0)
            throw new InvalidDataException("MHTML snapshot closing boundary does not match its parts.");
        HashSet<string> archivedLocations = ReadArchivedLocations(snapshot, boundary, closingStart);

        var additions = new StringBuilder();
        int added = 0;
        int addedBytes = 0;
        int alreadyArchived = 0;
        foreach (IResponse response in fonts) {
            if (!Uri.TryCreate(response.Url, UriKind.Absolute, out Uri? uri) || uri.Scheme != Uri.UriSchemeHttps)
                throw new InvalidDataException("A loaded font has no HTTPS resource URL.");
            string location = uri.AbsoluteUri;
            if (archivedLocations.Contains(location)) {
                alreadyArchived++;
                continue;
            }
            if (!response.Headers.TryGetValue("content-length", out string? lengthHeader)
                || !int.TryParse(lengthHeader, out int declaredBytes) || declaredBytes <= 0
                || declaredBytes > MaximumFontBytes || addedBytes + declaredBytes > MaximumTotalBytes)
                throw new InvalidDataException("Loaded font has no bounded Content-Length; refusing to buffer its response body.");
            byte[] body = await response.BodyAsync().ConfigureAwait(false);
            if (body.Length == 0 || body.Length > MaximumFontBytes || addedBytes + body.Length > MaximumTotalBytes)
                throw new InvalidDataException("Loaded font bytes exceed the bounded MHTML capture policy.");
            addedBytes += body.Length;
            string mediaType = response.Headers.TryGetValue("content-type", out string? contentType)
                ? contentType.Split(';', 2)[0].Trim() : "application/octet-stream";
            if (!IsSafeMediaType(mediaType)) mediaType = "application/octet-stream";
            additions.Append(boundary).Append("\r\nContent-Type: ").Append(mediaType)
                .Append("\r\nContent-Transfer-Encoding: base64\r\nContent-Location: ")
                .Append(location).Append("\r\n\r\n");
            string encoded = Convert.ToBase64String(body);
            for (int offset = 0; offset < encoded.Length; offset += 76)
                additions.Append(encoded, offset, Math.Min(76, encoded.Length - offset)).Append("\r\n");
            added++;
        }
        if (added == 0) return new Result(snapshot, fonts.Length, 0, 0, alreadyArchived);
        string archive = snapshot[..(closingStart + 2)] + additions.ToString() + snapshot[(closingStart + 2)..];
        return new Result(archive, fonts.Length, added, addedBytes, alreadyArchived);
    }

    private static HashSet<string> ReadArchivedLocations(string snapshot, string boundary, int closingStart) {
        var locations = new HashSet<string>(StringComparer.Ordinal);
        string delimiter = "\r\n" + boundary + "\r\n";
        int partStart = snapshot.IndexOf(delimiter, StringComparison.Ordinal);
        while (partStart >= 0 && partStart < closingStart) {
            int headerStart = partStart + delimiter.Length;
            int headerEnd = snapshot.IndexOf("\r\n\r\n", headerStart, StringComparison.Ordinal);
            if (headerEnd < 0 || headerEnd >= closingStart)
                throw new InvalidDataException("MHTML part has no terminating header block.");
            string? location = null;
            foreach (string line in snapshot[headerStart..headerEnd].Split("\r\n", StringSplitOptions.None)) {
                if (line.StartsWith("Content-Location:", StringComparison.OrdinalIgnoreCase)) {
                    AddLocation(location, locations);
                    location = line["Content-Location:".Length..].Trim();
                } else if (location != null && line.Length > 0 && (line[0] == ' ' || line[0] == '\t')) {
                    location += line.Trim();
                } else {
                    AddLocation(location, locations);
                    location = null;
                }
            }
            AddLocation(location, locations);
            partStart = snapshot.IndexOf(delimiter, headerEnd + 4, StringComparison.Ordinal);
        }
        return locations;
    }

    private static void AddLocation(string? value, HashSet<string> locations) {
        if (Uri.TryCreate(value, UriKind.Absolute, out Uri? uri)) locations.Add(uri.AbsoluteUri);
    }

    private static bool IsSafeMediaType(string value) =>
        value.Length > 0 && value.Length <= 100 && value.Contains('/')
        && value.All(character => char.IsAsciiLetterOrDigit(character)
            || character is '/' or '-' or '+' or '.');
}
