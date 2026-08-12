using OfficeIMO.Core.Internal;
using OfficeIMO.Provenance;

namespace OfficeIMO.Pdf;

/// <summary>Inspects and selectively removes the standards-defined PDF C2PA associated-file carrier.</summary>
public static class PdfProvenance {
    private const string C2paMimeType = "application/c2pa";

    /// <summary>Inspects a bounded PDF for C2PA Manifest Store associated files.</summary>
    public static OfficeProvenanceReport Inspect(
        byte[] pdf,
        OfficeProvenanceOptions? options = null,
        PdfReadOptions? readOptions = null) {
        Guard.NotNull(pdf, nameof(pdf));
        options ??= new OfficeProvenanceOptions();
        OfficeProvenanceBinary.ValidateLimits(options);
        if (pdf.LongLength > options.MaxAssetBytes) throw new InvalidDataException("The PDF exceeds the configured asset limit.");

        PdfReadDocument document = PdfReadDocument.Open(pdf, readOptions);
        IReadOnlyList<PdfExtractedAttachment> attachments = PdfAttachmentExtractor.ExtractAttachments(
            document,
            IsCandidate,
            Math.Min(options.MaxExpandedContainerBytes, MultiplySaturating(options.MaxManifestBytes, options.MaxCarriers)));
        HashSet<int> associatedFileSpecifications = CollectCatalogAssociatedFileSpecifications(document);
        var evidence = new List<OfficeProvenanceEvidence>();
        foreach (PdfExtractedAttachment attachment in attachments) {
            if (!IsCandidate(attachment)) continue;
            byte[] manifest = attachment.Bytes;
            if (manifest.LongLength > options.MaxManifestBytes) throw new InvalidDataException("A PDF provenance manifest exceeds the configured manifest limit.");
            bool valid = attachment.Relationship == PdfAssociatedFileRelationship.C2paManifest &&
                string.Equals(attachment.MimeType, C2paMimeType, StringComparison.OrdinalIgnoreCase) &&
                attachment.FileSpecObjectNumber > 0 &&
                associatedFileSpecifications.Contains(attachment.FileSpecObjectNumber) &&
                OfficeC2paManifestStore.IsValid(manifest, 0, manifest.Length, options.MaxManifestBytes, out _);
            if (evidence.Count >= options.MaxCarriers) throw new InvalidDataException($"The asset exceeds the configured carrier limit of {options.MaxCarriers}.");
            evidence.Add(new OfficeProvenanceEvidence(
                OfficeProvenanceCarrierKind.C2paManifest,
                $"PDF/Filespec[{attachment.FileSpecObjectNumber}]/{attachment.FileName}",
                valid,
                manifest.LongLength));
        }
        return new OfficeProvenanceReport(OfficeProvenanceAssetFormat.Pdf, evidence.AsReadOnly());
    }

    /// <summary>Inspects a bounded PDF file.</summary>
    public static OfficeProvenanceReport InspectFile(
        string filePath,
        OfficeProvenanceOptions? options = null,
        PdfReadOptions? readOptions = null) {
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("A file path is required.", nameof(filePath));
        options ??= new OfficeProvenanceOptions();
        byte[] pdf = ReadBounded(filePath, options.MaxAssetBytes);
        return Inspect(pdf, options, readOptions);
    }

    /// <summary>Removes selected structurally valid C2PA associated files through the proven PDF attachment rewrite.</summary>
    public static OfficeProvenanceRemovalResult Remove(
        byte[] pdf,
        OfficeProvenanceRemovalOptions? options = null,
        PdfReadOptions? readOptions = null) {
        Guard.NotNull(pdf, nameof(pdf));
        options ??= new OfficeProvenanceRemovalOptions();
        OfficeProvenanceReport before = Inspect(pdf, options.Limits, readOptions);
        if (!options.RemoveC2paManifests || before.Evidence.Count == 0) {
            return new OfficeProvenanceRemovalResult((byte[])pdf.Clone(), before, before, Array.Empty<OfficeProvenanceChange>(), false);
        }

        PdfReadDocument document = PdfReadDocument.Open(pdf, readOptions);
        IReadOnlyList<PdfExtractedAttachment> attachments = PdfAttachmentExtractor.ExtractAttachments(
            document,
            IsCandidate,
            Math.Min(options.Limits.MaxExpandedContainerBytes, MultiplySaturating(options.Limits.MaxManifestBytes, options.Limits.MaxCarriers)));
        IReadOnlyList<PdfAttachmentInfo> allAttachments = document.Attachments;
        var removeIndices = new List<int>();
        var usedAttachmentIndices = new HashSet<int>();
        var changes = new List<OfficeProvenanceChange>();
        int evidenceIndex = 0;
        for (int index = 0; index < attachments.Count; index++) {
            PdfExtractedAttachment attachment = attachments[index];
            if (!IsCandidate(attachment)) continue;
            OfficeProvenanceEvidence evidence = before.Evidence[evidenceIndex++];
            if (!evidence.IsStructurallyValid && options.RequireStructurallyValidCarrier) continue;
            int attachmentIndex = FindAttachmentIndex(allAttachments, attachment, usedAttachmentIndices);
            if (attachmentIndex < 0) throw new InvalidDataException("A PDF provenance attachment could not be matched to its descriptor.");
            removeIndices.Add(attachmentIndex);
            usedAttachmentIndices.Add(attachmentIndex);
            changes.Add(new OfficeProvenanceChange(
                OfficeProvenanceCarrierKind.C2paManifest,
                evidence.Location,
                removedBytes: 0));
        }
        if (removeIndices.Count == 0) {
            return new OfficeProvenanceRemovalResult((byte[])pdf.Clone(), before, before, Array.Empty<OfficeProvenanceChange>(), false);
        }

        PdfDocumentSecurityInfo security = PdfSyntax.ReadDocumentSecurityInfo(pdf, readOptions);
        if (security.HasSignatures) {
            string detail = options.SignatureMutationPolicy == OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
                ? "OfficeIMO.Pdf does not silently delete PDF signature revisions or fields; remove the signature through an explicit PDF signature workflow first."
                : "Removing provenance would invalidate PDF signatures.";
            throw new InvalidOperationException(detail);
        }

        PdfAttachmentEditResult edit = PdfAttachmentEditor.Edit(pdf, session => {
            for (int index = removeIndices.Count - 1; index >= 0; index--) session.RemoveAt(removeIndices[index]);
        }, readOptions, options.Limits.MaxExpandedContainerBytes);
        byte[] output = edit.ToBytes();
        PdfReadOptions outputReadOptions = PdfReadOptions.WithMinimumInputBytes(readOptions, output.LongLength);
        OfficeProvenanceReport after = Inspect(output, options.Limits, outputReadOptions);
        return new OfficeProvenanceRemovalResult(output, before, after, changes.AsReadOnly(), true);
    }

    /// <summary>Removes selected provenance and atomically writes the resulting PDF.</summary>
    public static OfficeProvenanceRemovalResult RemoveFile(
        string inputPath,
        string outputPath,
        OfficeProvenanceRemovalOptions? options = null,
        PdfReadOptions? readOptions = null) {
        if (string.IsNullOrWhiteSpace(inputPath)) throw new ArgumentException("An input path is required.", nameof(inputPath));
        if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("An output path is required.", nameof(outputPath));
        options ??= new OfficeProvenanceRemovalOptions();
        byte[] pdf = ReadBounded(inputPath, options.Limits.MaxAssetBytes);
        OfficeProvenanceRemovalResult result = Remove(pdf, options, readOptions);
        OfficeFileCommit.WriteAllBytes(Path.GetFullPath(outputPath), result.ToArray());
        return result;
    }

    private static bool IsCandidate(PdfExtractedAttachment attachment) =>
        attachment.Relationship == PdfAssociatedFileRelationship.C2paManifest ||
        string.Equals(attachment.MimeType, C2paMimeType, StringComparison.OrdinalIgnoreCase);

    private static bool IsCandidate(PdfAttachmentInfo attachment) =>
        attachment.Relationship == PdfAssociatedFileRelationship.C2paManifest ||
        string.Equals(attachment.MimeType, C2paMimeType, StringComparison.OrdinalIgnoreCase);

    private static HashSet<int> CollectCatalogAssociatedFileSpecifications(PdfReadDocument document) {
        var result = new HashSet<int>();
        PdfDictionary? catalog = PdfSyntax.FindCatalog(document.Objects, document.TrailerRaw);
        if (catalog == null || !catalog.Items.TryGetValue("AF", out PdfObject? associatedObject) ||
            PdfObjectLookup.Resolve(document.Objects, associatedObject) is not PdfArray associatedFiles) return result;
        foreach (PdfObject item in associatedFiles.Items) {
            if (item is PdfReference reference) result.Add(reference.ObjectNumber);
        }
        return result;
    }

    private static long MultiplySaturating(long value, int multiplier) =>
        value > long.MaxValue / multiplier ? long.MaxValue : value * multiplier;

    private static int FindAttachmentIndex(
        IReadOnlyList<PdfAttachmentInfo> attachments,
        PdfExtractedAttachment target,
        HashSet<int> usedIndices) {
        for (int index = 0; index < attachments.Count; index++) {
            if (usedIndices.Contains(index)) continue;
            PdfAttachmentInfo candidate = attachments[index];
            if (candidate.FileSpecObjectNumber == target.FileSpecObjectNumber &&
                candidate.EmbeddedFileObjectNumber == target.EmbeddedFileObjectNumber) return index;
        }
        return -1;
    }

    private static byte[] ReadBounded(string filePath, long maximumBytes) {
        using var stream = File.OpenRead(Path.GetFullPath(filePath));
        return OfficeProvenanceBinary.ReadBounded(stream, maximumBytes);
    }
}
