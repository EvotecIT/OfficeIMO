using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf;

/// <summary>
/// Represents a parsed PDF document with access to pages, catalog and metadata.
/// Note: MVP reader supports classic xref tables and simple stream parsing sufficient for OfficeIMO.Pdf output.
/// </summary>
public sealed class PdfReadDocument {
    private readonly Dictionary<int, PdfIndirectObject> _objects;
    private readonly string _trailerRaw;

    private PdfReadDocument(Dictionary<int, PdfIndirectObject> objects, string trailerRaw) {
        _objects = objects; _trailerRaw = trailerRaw;
        Pages = CollectPages();
    }

    /// <summary>All page objects discovered in document order.</summary>
    public IReadOnlyList<PdfReadPage> Pages { get; }

    /// <summary>Loads a PDF from bytes into a typed object model.</summary>
    public static PdfReadDocument Load(byte[] pdf) {
        var (map, trailer) = PdfSyntax.ParseObjects(pdf);
        return new PdfReadDocument(map, trailer);
    }

    /// <summary>Loads a PDF from a file path.</summary>
    public static PdfReadDocument Load(string path) => Load(File.ReadAllBytes(path));

    /// <summary>Returns document metadata (Title, Author, Subject, Keywords) when present.</summary>
    public (string? Title, string? Author, string? Subject, string? Keywords) GetMetadata() => PdfTextExtractor.GetMetadata(PdfEncoding.Latin1GetBytes(ToRaw()));

    private List<PdfReadPage> CollectPages() {
        var pages = new List<PdfReadPage>();
        // Heuristic: find all objects with /Type /Page
        foreach (var kv in _objects) {
            if (kv.Value.Value is PdfDictionary dict) {
                if (dict.Get<PdfName>("Type")?.Name == "Page") {
                    pages.Add(new PdfReadPage(kv.Key, dict, _objects));
                }
            }
        }
        // Sort by object number as a crude document order approximation
        pages.Sort((a, b) => a.ObjectNumber.CompareTo(b.ObjectNumber));
        return pages;
    }

    private string ToRaw() {
        // Reconstruct raw text for simple metadata extraction without reserialization; ok for small files.
        var sb = new StringBuilder();
        foreach (var kv in _objects.OrderBy(k => k.Key)) {
            sb.Append(kv.Key).Append(" 0 obj\n");
            if (kv.Value.Value is PdfStream s) {
                sb.Append("<< ");
                foreach (var d in s.Dictionary.Items) sb.Append('/').Append(d.Key).Append(' ').Append(' ').Append(' ');
                sb.Append(">>\nstream\n");
                sb.Append(PdfEncoding.Latin1GetString(s.Data)).Append("\nendstream\nendobj\n");
            } else {
                sb.Append("...\nendobj\n");
            }
        }
        sb.Append(_trailerRaw);
        return sb.ToString();
    }
}
