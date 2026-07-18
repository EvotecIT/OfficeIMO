using System.Security.Cryptography;

namespace OfficeIMO.Pdf;

/// <summary>Adds, replaces, renames, and removes embedded and associated files in existing PDFs.</summary>
internal static class PdfAttachmentEditor {
    /// <summary>Applies a collection edit through the shared full-rewrite planner and validates attachment readback.</summary>
    public static PdfAttachmentEditResult Edit(byte[] pdf, Action<PdfAttachmentEditSession> edit, PdfReadOptions? readOptions = null) {
        Guard.NotNull(pdf, nameof(pdf)); Guard.NotNull(edit, nameof(edit));
        PdfMutationPlan plan = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.ModifyAttachments, readOptions);
        IReadOnlyList<PdfExtractedAttachment> existing = PdfAttachmentExtractor.ExtractAttachments(PdfReadDocument.Open(pdf, readOptions));
        var session = new PdfAttachmentEditSession(existing.Select(static attachment => new PdfEmbeddedFile(attachment.FileName, attachment.Bytes, attachment.MimeType, attachment.Relationship, attachment.Description, attachment.CreationDate, attachment.ModificationDate)));
        edit(session);
        IReadOnlyList<PdfEmbeddedFile> target = session.Snapshot();
        byte[] output = PdfDocumentObjectGraphRewriter.Rewrite(pdf, readOptions, null, (objects, security) => {
            RewriteAttachmentGraph(objects, security, target);
            return security.InfoObjectNumber.HasValue && objects.ContainsKey(security.InfoObjectNumber.Value) ? security.InfoObjectNumber : null;
        });
        IReadOnlyList<PdfAttachmentValidation> validations = Validate(output, target);
        if (validations.Any(static validation => !validation.IsValid)) throw new InvalidOperationException("PDF attachment post-save validation failed; the artifact was not returned.");
        var preservationOptions = new PdfRewritePreservationOptions { OriginalReadOptions = readOptions, PreserveEmbeddedFiles = false, PreserveRevisionStructure = false };
        PdfRewritePreservationReport preservation = PdfRewritePreservation.AssertPreserved(pdf, output, preservationOptions);
        return new PdfAttachmentEditResult(output, plan, preservation, validations);
    }

    /// <summary>Adds one attachment.</summary>
    public static PdfAttachmentEditResult Add(byte[] pdf, PdfEmbeddedFile attachment, PdfReadOptions? readOptions = null) => Edit(pdf, session => session.Add(attachment), readOptions);
    /// <summary>Replaces one attachment by file name.</summary>
    public static PdfAttachmentEditResult Replace(byte[] pdf, string fileName, PdfEmbeddedFile replacement, PdfReadOptions? readOptions = null) => Edit(pdf, session => session.Replace(fileName, replacement), readOptions);
    /// <summary>Renames one attachment.</summary>
    public static PdfAttachmentEditResult Rename(byte[] pdf, string fileName, string newFileName, PdfReadOptions? readOptions = null) => Edit(pdf, session => session.Rename(fileName, newFileName), readOptions);
    /// <summary>Removes one attachment.</summary>
    public static PdfAttachmentEditResult Remove(byte[] pdf, string fileName, PdfReadOptions? readOptions = null) => Edit(pdf, session => session.Remove(fileName), readOptions);

    private static void RewriteAttachmentGraph(Dictionary<int, PdfIndirectObject> objects, PdfDocumentSecurityInfo security, IReadOnlyList<PdfEmbeddedFile> attachments) {
        if (!security.RootObjectNumber.HasValue || !objects.TryGetValue(security.RootObjectNumber.Value, out PdfIndirectObject? root) || root.Value is not PdfDictionary catalog) throw new InvalidOperationException("PDF catalog is not readable.");
        PdfDictionary? names = ResolveDictionary(objects, catalog.Items.TryGetValue("Names", out PdfObject? namesObject) ? namesObject : null);
        names?.Items.Remove("EmbeddedFiles");
        PdfAssociatedFileGraph.RemoveAssociatedFileReferences(objects);
        if (attachments.Count == 0) return;

        if (names == null) { names = new PdfDictionary(); catalog.Items["Names"] = names; }
        int nextObjectNumber = objects.Count == 0 ? 1 : objects.Keys.Max() + 1;
        var nameEntries = new List<(string Name, PdfReference FileSpec)>(attachments.Count);
        var associated = new PdfArray();
        foreach (PdfEmbeddedFile attachment in attachments.OrderBy(static file => file.FileName, StringComparer.Ordinal)) {
            byte[] data = attachment.DataSnapshot;
            int streamNumber = nextObjectNumber++;
            int fileSpecNumber = nextObjectNumber++;
            PdfDictionary streamDictionary = BuildEmbeddedFileDictionary(attachment, data);
            objects[streamNumber] = new PdfIndirectObject(streamNumber, 0, new PdfStream(streamDictionary, data));
            PdfDictionary fileSpec = BuildFileSpec(attachment, streamNumber);
            objects[fileSpecNumber] = new PdfIndirectObject(fileSpecNumber, 0, fileSpec);
            var reference = new PdfReference(fileSpecNumber, 0);
            nameEntries.Add((attachment.FileName, reference));
            if (attachment.Relationship != PdfAssociatedFileRelationship.Unspecified) associated.Items.Add(reference);
        }
        var nameArray = new PdfArray();
        foreach ((string name, PdfReference reference) in nameEntries) { nameArray.Items.Add(new PdfStringObj(name, true)); nameArray.Items.Add(reference); }
        var embeddedFiles = new PdfDictionary(); embeddedFiles.Items["Names"] = nameArray; names.Items["EmbeddedFiles"] = embeddedFiles;
        if (associated.Items.Count > 0) catalog.Items["AF"] = associated;
    }

    private static PdfDictionary BuildEmbeddedFileDictionary(PdfEmbeddedFile attachment, byte[] data) {
        var dictionary = new PdfDictionary(); dictionary.Items["Type"] = new PdfName("EmbeddedFile"); dictionary.Items["Length"] = new PdfNumber(data.Length);
        if (!string.IsNullOrWhiteSpace(attachment.MimeType)) dictionary.Items["Subtype"] = new PdfName(attachment.MimeType!);
        var parameters = new PdfDictionary(); parameters.Items["Size"] = new PdfNumber(data.Length); parameters.Items["CheckSum"] = new PdfStringObj(ComputeChecksum(data));
        if (attachment.CreationDate.HasValue) parameters.Items["CreationDate"] = new PdfStringObj(FormatPdfDate(attachment.CreationDate.Value));
        if (attachment.ModificationDate.HasValue) parameters.Items["ModDate"] = new PdfStringObj(FormatPdfDate(attachment.ModificationDate.Value));
        dictionary.Items["Params"] = parameters;
        return dictionary;
    }

    private static PdfDictionary BuildFileSpec(PdfEmbeddedFile attachment, int streamNumber) {
        var dictionary = new PdfDictionary(); dictionary.Items["Type"] = new PdfName("Filespec"); dictionary.Items["F"] = new PdfStringObj(attachment.FileName, true); dictionary.Items["UF"] = new PdfStringObj(attachment.FileName, true);
        var embedded = new PdfDictionary(); var reference = new PdfReference(streamNumber, 0); embedded.Items["F"] = reference; embedded.Items["UF"] = reference; dictionary.Items["EF"] = embedded;
        dictionary.Items["AFRelationship"] = new PdfName(PdfEmbeddedFileDictionaryBuilder.GetRelationshipName(attachment.Relationship));
        if (attachment.Description != null) dictionary.Items["Desc"] = new PdfStringObj(attachment.Description, true);
        return dictionary;
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<PdfAttachmentValidation> Validate(byte[] pdf, IReadOnlyList<PdfEmbeddedFile> target) {
        IReadOnlyList<PdfExtractedAttachment> actual = PdfAttachmentExtractor.ExtractAttachments(pdf);
        var result = new List<PdfAttachmentValidation>(target.Count);
        foreach (PdfEmbeddedFile expected in target) {
            PdfExtractedAttachment? found = actual.FirstOrDefault(item => string.Equals(item.FileName, expected.FileName, StringComparison.Ordinal));
            byte[] data = expected.DataSnapshot;
            bool payload = found != null && found.Bytes.SequenceEqual(data);
            bool metadata = found != null && string.Equals(found.MimeType, expected.MimeType, StringComparison.Ordinal) && string.Equals(found.Description, expected.Description, StringComparison.Ordinal) && found.Relationship == expected.Relationship && DatesMatch(found.CreationDate, expected.CreationDate) && DatesMatch(found.ModificationDate, expected.ModificationDate);
            result.Add(new PdfAttachmentValidation(expected.FileName, ToHex(ComputeChecksum(data)), payload, metadata));
        }
        if (actual.Count != target.Count) throw new InvalidOperationException("PDF attachment post-save validation found an unexpected attachment count.");
        return result.AsReadOnly();
    }

    private static PdfDictionary? ResolveDictionary(Dictionary<int, PdfIndirectObject> objects, PdfObject? value) => PdfObjectLookup.Resolve(objects, value) as PdfDictionary;
    private static bool DatesMatch(DateTimeOffset? actual, DateTimeOffset? expected) => !actual.HasValue && !expected.HasValue || actual.HasValue && expected.HasValue && actual.Value.ToUnixTimeSeconds() == expected.Value.ToUnixTimeSeconds() && actual.Value.Offset == expected.Value.Offset;

    private static string FormatPdfDate(DateTimeOffset value) {
        string sign = value.Offset < TimeSpan.Zero ? "-" : "+";
        TimeSpan offset = value.Offset.Duration();
        return "D:" + value.ToString("yyyyMMddHHmmss", System.Globalization.CultureInfo.InvariantCulture) + sign + offset.Hours.ToString("00", System.Globalization.CultureInfo.InvariantCulture) + "'" + offset.Minutes.ToString("00", System.Globalization.CultureInfo.InvariantCulture) + "'";
    }

    #pragma warning disable CA5351, CA1850
    private static byte[] ComputeChecksum(byte[] data) { using (MD5 md5 = MD5.Create()) return md5.ComputeHash(data); }
    #pragma warning restore CA5351, CA1850
    private static string ToHex(byte[] bytes) { var builder = new System.Text.StringBuilder(bytes.Length * 2); for (int i = 0; i < bytes.Length; i++) builder.Append(bytes[i].ToString("x2", System.Globalization.CultureInfo.InvariantCulture)); return builder.ToString(); }
}
