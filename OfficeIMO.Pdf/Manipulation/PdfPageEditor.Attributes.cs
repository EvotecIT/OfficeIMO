using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageEditor {
    /// <summary>Writes page-local rotation and box overrides through the route proven by the shared planner.</summary>
    private static byte[] WritePageAttributes(
        byte[] pdf,
        PdfReadDocument document,
        Dictionary<int, Dictionary<string, PdfObject>> overrides,
        PdfMutationPlan plan,
        PdfLoadOptions? readOptions,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (plan.ExecutionMode == PdfMutationExecutionMode.FullRewrite) {
            Dictionary<int, PdfIndirectObject> objects = document.Objects;
            return PdfPageExtractor.ExtractPages(
                objects,
                document.UncheckedMetadata,
                document.Pages.Select(static page => page.ObjectNumber).ToArray(),
                overrides,
                catalogState: PdfPageExtractor.ExtractCatalogRewriteState(objects, document.TrailerRaw, document, cancellationToken),
                fileVersion: PdfPageExtractor.GetSourceFileVersion(pdf),
                cancellationToken: cancellationToken);
        }

        if (overrides.Count == 0) return (byte[])pdf.Clone();
        // ReadDocument resolves inherited values and may repair semantic links in memory.
        // Append from the original parsed dictionaries so unrelated repairs are not persisted.
        var (originalObjects, trailerRaw) = PdfSyntax.ParseObjects(pdf, readOptions, out _, out _, cancellationToken);
        foreach (KeyValuePair<int, Dictionary<string, PdfObject>> item in overrides) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfIndirectObject original = originalObjects[item.Key];
            if (original.Value is not PdfDictionary page) {
                throw new InvalidOperationException("The selected page does not resolve to a page dictionary.");
            }
            var updated = new PdfDictionary();
            foreach (KeyValuePair<string, PdfObject> entry in page.Items) updated.Items.Add(entry.Key, entry.Value);
            foreach (KeyValuePair<string, PdfObject> entry in item.Value) {
                // This route must never acquire membership, content, resource, or tag edits.
                if (entry.Key != "Rotate" && entry.Key != "MediaBox" && entry.Key != "CropBox" &&
                    entry.Key != "TrimBox" && entry.Key != "BleedBox" && entry.Key != "ArtBox") {
                    throw new InvalidOperationException("Only rotation and boundary boxes can use a page-attribute update.");
                }
                updated.Items[entry.Key] = entry.Value;
            }
            originalObjects[item.Key] = new PdfIndirectObject(original.ObjectNumber, original.Generation, updated);
        }

        PdfDocumentSecurityInfo security = plan.Preflight.Probe.Security;
        PdfStandardSecurityHandler? encryptionHandler = null;
        if (security.HasEncryption &&
            !PdfSyntax.TryCreateDecryptor(originalObjects, trailerRaw, readOptions, out encryptionHandler, cancellationToken)) {
            throw new PdfUnsupportedEncryptionException("PDF encryption context could not be created for the page-attribute update.");
        }
        byte[] output = PdfIncrementalObjectWriter.Append(
            pdf, originalObjects, security, trailerRaw, overrides.Keys,
            encryptionHandler: encryptionHandler, cancellationToken: cancellationToken);
        PdfLoadOptions outputOptions = PdfLoadOptions.ForGeneratedOutput(
            readOptions, pdf, output, new PdfGeneratedOutputGrowth(additionalRevisions: 1));
        PdfSignatureMutationReport proof = PdfSignatureMutationAnalyzer.Analyze(
            pdf, output, PdfMutationOperation.ModifyPageAttributes,
            readOptions: readOptions, afterReadOptions: outputOptions);
        if (!proof.IsPreservedAppendOnlyMutation) {
            throw new InvalidOperationException("The page-attribute update did not preserve the input prefix, revision chain, and existing signature byte ranges.");
        }
        cancellationToken.ThrowIfCancellationRequested();
        return output;
    }
}
