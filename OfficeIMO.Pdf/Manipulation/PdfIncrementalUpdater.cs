using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>
/// Append-only PDF update helpers for changes that can be represented as a new incremental revision.
/// </summary>
public static partial class PdfIncrementalUpdater {
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
        var commonBlockers = new List<string>();
        var metadataBlockers = new List<string>();
        var formBlockers = new List<string>();
        var warnings = new List<string>();
        if (security.HasEncryption) {
            commonBlockers.Add("Encrypted");
        }

        bool hasSignatureContent = security.SignatureFieldCount > 0 || security.SignatureCount > 0 || security.HasByteRange;
        if (security.HasUsageRights) {
            commonBlockers.Add("UsageRights");
        }

        if (security.HasXrefStreams) {
            commonBlockers.Add("XrefStream");
        }

        if (!security.RootObjectNumber.HasValue) {
            commonBlockers.Add("MissingRoot");
        }

        if (!security.LastStartXrefOffset.HasValue) {
            commonBlockers.Add("MissingStartXref");
        }

        metadataBlockers.AddRange(commonBlockers);
        formBlockers.AddRange(commonBlockers);

        if (hasSignatureContent) {
            metadataBlockers.Add("Signed");
            if (!CanAppendFormFieldsWithDocMDP(security, null)) {
                formBlockers.Add("Signed");
            } else {
                warnings.Add("SignedDocMDPFormFill");
            }
        }

        if (security.HasDocMDPPermissions) {
            metadataBlockers.Add("DocMDP");
            if (!CanAppendFormFieldsWithDocMDP(security, null)) {
                formBlockers.Add("DocMDP");
            } else {
                warnings.Add("DocMDPAllowsFormFill");
            }
        }

        if (security.HasIncrementalUpdates) {
            warnings.Add("ExistingIncrementalRevisions");
        }

        if (security.AcroFormAppendOnly) {
            warnings.Add("AcroFormAppendOnly");
        }

        var supported = new List<string>();
        if (metadataBlockers.Count == 0) {
            supported.Add("Metadata");
        }

        if (formBlockers.Count == 0) {
            supported.Add("FormFill");
        }

        bool canPrepareSignature =
            commonBlockers.Count == 0 &&
            !hasSignatureContent &&
            !security.HasDocMDPPermissions;
        if (canPrepareSignature) {
            supported.Add("SignaturePrepare");
        }

        var blocked = new List<string>();
        if (metadataBlockers.Count > 0) {
            blocked.Add("Metadata");
        }

        if (formBlockers.Count > 0) {
            blocked.Add("FormFill");
        }

        if (!canPrepareSignature) {
            blocked.Add("SignaturePrepare");
        }

        blocked.Add("Annotations");
        blocked.Add("PageTree");
        blocked.Add("Attachments");

        var blockers = metadataBlockers
            .Concat(formBlockers)
            .Distinct(StringComparer.Ordinal)
            .ToArray();

        return new PdfAppendOnlyMutationReport(
            security,
            supported.AsReadOnly(),
            blocked.AsReadOnly(),
            blockers,
            warnings.Distinct(StringComparer.Ordinal).ToArray());
    }

    private static bool CanAppendFormFieldsWithDocMDP(PdfDocumentSecurityInfo security, IEnumerable<string>? fieldNames) {
        if (!security.HasDocMDPPermissions ||
            !security.DocMDPPermissionLevel.HasValue ||
            security.DocMDPPermissionLevel.Value < 2 ||
            security.DocMDPPermissionLevel.Value > 3) {
            return false;
        }

        if (fieldNames is null) {
            return !security.Signatures.Any(static signature => LocksEveryField(signature.FieldLock));
        }

        return GetFirstLockedFormFieldName(security, fieldNames) is null;
    }

    private static string? GetFirstLockedFormFieldName(PdfDocumentSecurityInfo security, IEnumerable<string> fieldNames) {
        var requested = new HashSet<string>(fieldNames, StringComparer.Ordinal);
        if (requested.Count == 0) {
            return null;
        }

        foreach (PdfSignatureInfo signature in security.Signatures) {
            PdfSignatureFieldLockInfo? fieldLock = signature.FieldLock;
            if (fieldLock is null) {
                continue;
            }

            foreach (string fieldName in requested) {
                if (IsFieldLocked(fieldLock, fieldName)) {
                    return fieldName;
                }
            }
        }

        return null;
    }

    private static bool IsFieldLocked(PdfSignatureFieldLockInfo fieldLock, string fieldName) {
        if (fieldLock.LocksAllFields) {
            return true;
        }

        bool listed = fieldLock.Fields.Contains(fieldName, StringComparer.Ordinal);
        if (fieldLock.LocksIncludedFields) {
            return listed;
        }

        if (fieldLock.LocksAllExceptListedFields) {
            return !listed;
        }

        return false;
    }

    private static bool LocksEveryField(PdfSignatureFieldLockInfo? fieldLock) {
        return fieldLock is not null &&
            (fieldLock.LocksAllFields ||
            (fieldLock.LocksAllExceptListedFields && fieldLock.Fields.Count == 0));
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
