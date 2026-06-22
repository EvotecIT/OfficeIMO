using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>
/// Append-only PDF update helpers for changes that can be represented as a new incremental revision.
/// </summary>
public static class PdfIncrementalUpdater {
    /// <summary>
    /// Analyzes append-only mutation support for a PDF byte array.
    /// </summary>
    public static PdfAppendOnlyMutationReport AnalyzeAppendOnlyMutation(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));
        return BuildAppendOnlyMutationReport(PdfSyntax.ReadDocumentSecurityInfo(pdf));
    }

    /// <summary>Analyzes append-only mutation support for a readable PDF stream.</summary>
    public static PdfAppendOnlyMutationReport AnalyzeAppendOnlyMutation(Stream input) {
        Guard.NotNull(input, nameof(input));
        if (!input.CanRead) {
            throw new ArgumentException("Stream must be readable.", nameof(input));
        }

        using var buffer = new MemoryStream();
        input.CopyTo(buffer);
        return AnalyzeAppendOnlyMutation(buffer.ToArray());
    }

    /// <summary>Analyzes append-only mutation support for a PDF file.</summary>
    public static PdfAppendOnlyMutationReport AnalyzeAppendOnlyMutation(string inputPath) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return AnalyzeAppendOnlyMutation(File.ReadAllBytes(inputPath));
    }

    /// <summary>
    /// Appends a metadata-only revision to a PDF byte array without rewriting the existing bytes.
    /// </summary>
    public static byte[] UpdateMetadata(byte[] pdf, string? title = null, string? author = null, string? subject = null, string? keywords = null) {
        Guard.NotNull(pdf, nameof(pdf));

        PdfDocumentSecurityInfo security = PdfSyntax.ReadDocumentSecurityInfo(pdf);
        ValidateAppendOnlyMetadataInput(security);

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf);
        if (!security.RootObjectNumber.HasValue) {
            throw new InvalidOperationException("PDF root catalog reference is required for an incremental metadata update.");
        }

        if (!security.LastStartXrefOffset.HasValue) {
            throw new InvalidOperationException("PDF startxref offset is required for an incremental metadata update.");
        }

        PdfMetadata existing = PdfInspector.Inspect(pdf).Metadata;
        var updated = new PdfMetadata {
            Title = title ?? existing.Title,
            Author = author ?? existing.Author,
            Subject = subject ?? existing.Subject,
            Keywords = keywords ?? existing.Keywords
        };

        int newInfoObjectNumber = objects.Count == 0 ? 1 : objects.Keys.Max() + 1;
        int size = newInfoObjectNumber + 1;
        byte[] infoObject = PdfObjectBytes.WrapIndirectObject(newInfoObjectNumber, PdfInfoDictionaryBuilder.Build(updated));

        using var output = new MemoryStream(pdf.Length + infoObject.Length + 256);
        output.Write(pdf, 0, pdf.Length);
        if (pdf.Length == 0 || (pdf[pdf.Length - 1] != (byte)'\n' && pdf[pdf.Length - 1] != (byte)'\r')) {
            output.WriteByte((byte)'\n');
        }

        long objectOffset = output.Position;
        output.Write(infoObject, 0, infoObject.Length);
        long xrefOffset = output.Position;

        using var writer = new StreamWriter(output, Encoding.ASCII, 1024, leaveOpen: true) { NewLine = "\n" };
        writer.WriteLine("xref");
        writer.WriteLine(newInfoObjectNumber.ToString(CultureInfo.InvariantCulture) + " 1");
        writer.WriteLine(objectOffset.ToString("0000000000", CultureInfo.InvariantCulture) + " 00000 n ");
        writer.WriteLine("trailer");
        writer.WriteLine("<< /Size " + size.ToString(CultureInfo.InvariantCulture) +
            " /Root " + PdfSyntaxEscaper.IndirectReference(security.RootObjectNumber.Value) +
            " /Info " + PdfSyntaxEscaper.IndirectReference(newInfoObjectNumber) +
            " /Prev " + security.LastStartXrefOffset.Value.ToString(CultureInfo.InvariantCulture) +
            ReadTrailerIdEntry(trailerRaw) +
            " >>");
        writer.WriteLine("startxref");
        writer.WriteLine(xrefOffset.ToString(CultureInfo.InvariantCulture));
        writer.WriteLine("%%EOF");
        writer.Flush();

        return output.ToArray();
    }

    /// <summary>Appends a metadata-only revision to a PDF stream.</summary>
    public static byte[] UpdateMetadata(Stream input, string? title = null, string? author = null, string? subject = null, string? keywords = null) {
        Guard.NotNull(input, nameof(input));
        if (!input.CanRead) {
            throw new ArgumentException("Stream must be readable.", nameof(input));
        }

        using var buffer = new MemoryStream();
        input.CopyTo(buffer);
        return UpdateMetadata(buffer.ToArray(), title, author, subject, keywords);
    }

    /// <summary>Appends a metadata-only revision to a PDF file and writes the result to <paramref name="outputPath"/>.</summary>
    public static void UpdateMetadata(string inputPath, string outputPath, string? title = null, string? author = null, string? subject = null, string? keywords = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNullOrWhiteSpace(outputPath, nameof(outputPath));
        File.WriteAllBytes(outputPath, UpdateMetadata(File.ReadAllBytes(inputPath), title, author, subject, keywords));
    }

    private static void ValidateAppendOnlyMetadataInput(PdfDocumentSecurityInfo security) {
        PdfAppendOnlyMutationReport report = BuildAppendOnlyMutationReport(security);
        if (!report.CanAppendMetadata) {
            throw new NotSupportedException("Incremental metadata updates are not supported for this PDF: " + string.Join(", ", report.Blockers));
        }
    }

    private static PdfAppendOnlyMutationReport BuildAppendOnlyMutationReport(PdfDocumentSecurityInfo security) {
        var blockers = new List<string>();
        var warnings = new List<string>();
        if (security.HasEncryption) {
            blockers.Add("Encrypted");
        }

        if (security.HasSignatures) {
            blockers.Add("Signed");
        }

        if (security.HasDocMDPPermissions) {
            blockers.Add("DocMDP");
        }

        if (security.HasUsageRights) {
            blockers.Add("UsageRights");
        }

        if (security.HasXrefStreams) {
            blockers.Add("XrefStream");
        }

        if (!security.RootObjectNumber.HasValue) {
            blockers.Add("MissingRoot");
        }

        if (!security.LastStartXrefOffset.HasValue) {
            blockers.Add("MissingStartXref");
        }

        if (security.HasIncrementalUpdates) {
            warnings.Add("ExistingIncrementalRevisions");
        }

        if (security.AcroFormAppendOnly) {
            warnings.Add("AcroFormAppendOnly");
        }

        string[] supported = blockers.Count == 0 ? new[] { "Metadata" } : Array.Empty<string>();
        var blocked = new List<string>();
        if (blockers.Count > 0) {
            blocked.Add("Metadata");
        }

        blocked.Add("FormFill");
        blocked.Add("Annotations");
        blocked.Add("PageTree");
        blocked.Add("Attachments");

        return new PdfAppendOnlyMutationReport(
            security,
            supported,
            blocked.AsReadOnly(),
            blockers.AsReadOnly(),
            warnings.AsReadOnly());
    }

    private static string ReadTrailerIdEntry(string trailerRaw) {
        int nameIndex = IndexOfName(trailerRaw, "ID");
        if (nameIndex < 0) {
            return string.Empty;
        }

        int start = trailerRaw.IndexOf('[', nameIndex);
        if (start < 0) {
            return string.Empty;
        }

        int depth = 0;
        for (int i = start; i < trailerRaw.Length; i++) {
            if (trailerRaw[i] == '[') {
                depth++;
            } else if (trailerRaw[i] == ']') {
                depth--;
                if (depth == 0) {
                    return " /ID " + trailerRaw.Substring(start, i - start + 1).Trim();
                }
            }
        }

        return string.Empty;
    }

    private static int IndexOfName(string value, string name) {
        string token = "/" + name;
        int index = 0;
        while (index < value.Length) {
            int found = value.IndexOf(token, index, StringComparison.Ordinal);
            if (found < 0) {
                return -1;
            }

            int after = found + token.Length;
            if (after >= value.Length || IsDelimiter(value[after])) {
                return found;
            }

            index = after;
        }

        return -1;
    }

    private static bool IsDelimiter(char value) =>
        char.IsWhiteSpace(value) ||
        value == '/' ||
        value == '<' ||
        value == '>' ||
        value == '[' ||
        value == ']' ||
        value == '(' ||
        value == ')';
}
