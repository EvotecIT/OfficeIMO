using System.Security.Cryptography;

namespace OfficeIMO.Pdf;

/// <summary>Creates append-only DSS/VRI revisions for cryptographically verified PDF signatures.</summary>
public static class PdfLongTermValidationEnricher {
    /// <summary>
    /// Appends DER-encoded validation material for one signature without changing any existing PDF byte.
    /// Signature math and the signed-content digest must first validate through <paramref name="cryptographyProvider"/>.
    /// </summary>
    public static PdfLongTermValidationEnrichmentResult Enrich(
        byte[] pdf,
        PdfLongTermValidationEvidence evidence,
        IPdfSignatureCryptographyProvider cryptographyProvider,
        PdfReadOptions? readOptions = null) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(evidence, nameof(evidence));
        Guard.NotNull(cryptographyProvider, nameof(cryptographyProvider));

        PdfSignatureValidationReport before = PdfSignatureValidator.Validate(pdf, cryptographyProvider, readOptions);
        PdfSignatureValidationResult target = FindVerifiedTarget(before, evidence.SignatureObjectNumber);
        PdfMutationPlan plan = PdfMutationPlanner.Plan(pdf, PdfMutationOperation.EnrichLongTermValidation, readOptions);
        if (!plan.CanExecute) {
            throw new NotSupportedException("DSS/VRI enrichment is not supported for this PDF: " + string.Join(", ", plan.BlockerCodes));
        }

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf, readOptions);
        PdfDocumentSecurityInfo security = before.Security;
        if (!security.RootObjectNumber.HasValue ||
            !objects.TryGetValue(security.RootObjectNumber.Value, out PdfIndirectObject? rootObject) ||
            rootObject.Value is not PdfDictionary catalog) {
            throw new InvalidOperationException("PDF root catalog dictionary is required for DSS/VRI enrichment.");
        }

        int nextObjectNumber = objects.Count == 0 ? 1 : objects.Keys.Max() + 1;
        var changedObjects = new HashSet<int>();
        IReadOnlyList<int> certificateObjects = AddEvidenceStreams(objects, evidence.CertificateValues, ref nextObjectNumber, changedObjects);
        IReadOnlyList<int> ocspObjects = AddEvidenceStreams(objects, evidence.OcspValues, ref nextObjectNumber, changedObjects);
        IReadOnlyList<int> crlObjects = AddEvidenceStreams(objects, evidence.CrlValues, ref nextObjectNumber, changedObjects);

        int dssObjectNumber = EnsureDssDictionary(objects, catalog, ref nextObjectNumber, changedObjects, out PdfDictionary dss);
        dss.Items["Type"] = new PdfName("DSS");
        AppendTopLevelReferences(objects, dss, "Certs", certificateObjects);
        AppendTopLevelReferences(objects, dss, "OCSPs", ocspObjects);
        AppendTopLevelReferences(objects, dss, "CRLs", crlObjects);

        string vriKey = ComputeVriKey(pdf, target);
        PdfDictionary vri = CloneResolvedDictionary(objects, dss.Items.TryGetValue("VRI", out PdfObject? currentVri) ? currentVri : null);
        PdfDictionary vriEntry = CloneResolvedDictionary(objects, vri.Items.TryGetValue(vriKey, out PdfObject? currentVriEntry) ? currentVriEntry : null);
        vriEntry.Items["Type"] = new PdfName("VRI");
        AppendReferenceArray(objects, vriEntry, "Cert", certificateObjects);
        AppendReferenceArray(objects, vriEntry, "OCSP", ocspObjects);
        AppendReferenceArray(objects, vriEntry, "CRL", crlObjects);
        int vriEntryObjectNumber = nextObjectNumber++;
        objects[vriEntryObjectNumber] = new PdfIndirectObject(vriEntryObjectNumber, 0, vriEntry);
        changedObjects.Add(vriEntryObjectNumber);

        vri.Items[vriKey] = new PdfReference(vriEntryObjectNumber, 0);
        dss.Items["VRI"] = vri;
        changedObjects.Add(dssObjectNumber);
        AddEtsiExtension(objects, catalog);
        changedObjects.Add(security.RootObjectNumber.Value);

        byte[] enriched = PdfIncrementalObjectWriter.Append(
            pdf,
            objects,
            security,
            trailerRaw,
            changedObjects);
        PdfSignatureValidationReport after = PdfSignatureValidator.Validate(enriched, cryptographyProvider, readOptions);
        PdfSignatureMutationReport mutation = PdfSignatureMutationAnalyzer.Analyze(
            pdf,
            enriched,
            PdfMutationOperation.EnrichLongTermValidation,
            readOptions: readOptions);
        var result = new PdfLongTermValidationEnrichmentResult(
            enriched,
            vriKey,
            evidence,
            before,
            after,
            mutation,
            certificateObjects,
            ocspObjects,
            crlObjects);
        PdfSignatureValidationResult targetAfter = FindVerifiedTarget(after, evidence.SignatureObjectNumber);
        _ = targetAfter;
        if (!result.IsVerifiedAppendOnlyEnrichment) {
            throw new InvalidOperationException("DSS/VRI enrichment did not pass append-only signature and evidence readback proofs.");
        }

        return result;
    }

    /// <summary>Enriches a readable PDF stream.</summary>
    public static PdfLongTermValidationEnrichmentResult Enrich(
        Stream input,
        PdfLongTermValidationEvidence evidence,
        IPdfSignatureCryptographyProvider cryptographyProvider,
        PdfReadOptions? readOptions = null) {
        Guard.NotNull(input, nameof(input));
        if (!input.CanRead) {
            throw new ArgumentException("Stream must be readable.", nameof(input));
        }

        using var buffer = new MemoryStream();
        input.CopyTo(buffer);
        return Enrich(buffer.ToArray(), evidence, cryptographyProvider, readOptions);
    }

    /// <summary>Enriches a PDF file and writes the verified append-only result.</summary>
    public static PdfLongTermValidationEnrichmentResult Enrich(
        string inputPath,
        string outputPath,
        PdfLongTermValidationEvidence evidence,
        IPdfSignatureCryptographyProvider cryptographyProvider,
        PdfReadOptions? readOptions = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNullOrWhiteSpace(outputPath, nameof(outputPath));
        PdfLongTermValidationEnrichmentResult result = Enrich(File.ReadAllBytes(inputPath), evidence, cryptographyProvider, readOptions);
        File.WriteAllBytes(outputPath, result.Pdf);
        return result;
    }

    private static PdfSignatureValidationResult FindVerifiedTarget(PdfSignatureValidationReport report, int objectNumber) {
        PdfSignatureValidationResult? target = report.Signatures.FirstOrDefault(signature => signature.Signature.ObjectNumber == objectNumber);
        if (target is null) {
            throw new ArgumentException("PDF does not contain signature object " + objectNumber + ".", nameof(objectNumber));
        }

        if (!target.IsStructurallyValid) {
            throw new InvalidOperationException("DSS/VRI evidence cannot be attached to a structurally invalid signature.");
        }

        if (!target.Signature.HasRecognizedSubFilter) {
            throw new NotSupportedException("DSS/VRI enrichment currently supports CMS, CAdES, and RFC 3161 PDF signatures.");
        }

        if (target.CryptographicResult is null || !target.CryptographicResult.IsMathematicallyValid) {
            throw new InvalidOperationException("DSS/VRI evidence requires valid signature math and signed-content digest verification.");
        }

        return target;
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<int> AddEvidenceStreams(
        Dictionary<int, PdfIndirectObject> objects,
        IReadOnlyList<byte[]> values,
        ref int nextObjectNumber,
        HashSet<int> changedObjects) {
        var objectNumbers = new List<int>(values.Count);
        for (int i = 0; i < values.Count; i++) {
            int objectNumber = nextObjectNumber++;
            objects[objectNumber] = new PdfIndirectObject(
                objectNumber,
                0,
                new PdfStream(new PdfDictionary(), (byte[])values[i].Clone()));
            changedObjects.Add(objectNumber);
            objectNumbers.Add(objectNumber);
        }

        return objectNumbers.AsReadOnly();
    }

    private static int EnsureDssDictionary(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary catalog,
        ref int nextObjectNumber,
        HashSet<int> changedObjects,
        out PdfDictionary dss) {
        if (catalog.Items.TryGetValue("DSS", out PdfObject? current)) {
            if (current is PdfReference reference && PdfObjectLookup.Resolve(objects, reference) is PdfDictionary referenced) {
                dss = referenced;
                return reference.ObjectNumber;
            }

            if (PdfObjectLookup.Resolve(objects, current) is PdfDictionary direct) {
                int materializedObjectNumber = nextObjectNumber++;
                dss = direct;
                objects[materializedObjectNumber] = new PdfIndirectObject(materializedObjectNumber, 0, dss);
                catalog.Items["DSS"] = new PdfReference(materializedObjectNumber, 0);
                changedObjects.Add(materializedObjectNumber);
                return materializedObjectNumber;
            }
        }

        int objectNumber = nextObjectNumber++;
        dss = new PdfDictionary();
        objects[objectNumber] = new PdfIndirectObject(objectNumber, 0, dss);
        catalog.Items["DSS"] = new PdfReference(objectNumber, 0);
        changedObjects.Add(objectNumber);
        return objectNumber;
    }

    private static void AppendTopLevelReferences(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary dss,
        string key,
        IReadOnlyList<int> objectNumbers) {
        if (objectNumbers.Count == 0) {
            return;
        }

        var combined = new PdfArray();
        if (dss.Items.TryGetValue(key, out PdfObject? current) && PdfObjectLookup.Resolve(objects, current) is PdfArray existing) {
            combined.Items.AddRange(existing.Items);
        }

        for (int i = 0; i < objectNumbers.Count; i++) {
            combined.Items.Add(new PdfReference(objectNumbers[i], 0));
        }

        dss.Items[key] = combined;
    }

    private static void AppendReferenceArray(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary dictionary,
        string key,
        IReadOnlyList<int> objectNumbers) {
        if (objectNumbers.Count == 0) {
            return;
        }

        var array = new PdfArray();
        if (dictionary.Items.TryGetValue(key, out PdfObject? current) && PdfObjectLookup.Resolve(objects, current) is PdfArray existing) {
            array.Items.AddRange(existing.Items);
        }

        for (int i = 0; i < objectNumbers.Count; i++) {
            array.Items.Add(new PdfReference(objectNumbers[i], 0));
        }

        dictionary.Items[key] = array;
    }

    private static PdfDictionary CloneResolvedDictionary(Dictionary<int, PdfIndirectObject> objects, PdfObject? value) {
        var result = new PdfDictionary();
        if (PdfObjectLookup.Resolve(objects, value) is PdfDictionary existing) {
            foreach (KeyValuePair<string, PdfObject> item in existing.Items) {
                result.Items[item.Key] = item.Value;
            }
        }

        return result;
    }

    private static void AddEtsiExtension(Dictionary<int, PdfIndirectObject> objects, PdfDictionary catalog) {
        PdfDictionary extensions = CloneResolvedDictionary(
            objects,
            catalog.Items.TryGetValue("Extensions", out PdfObject? current) ? current : null);
        if (!extensions.Items.ContainsKey("ESIC")) {
            var esic = new PdfDictionary();
            esic.Items["BaseVersion"] = new PdfName("1.7");
            esic.Items["ExtensionLevel"] = new PdfNumber(1);
            extensions.Items["ESIC"] = esic;
        }

        catalog.Items["Extensions"] = extensions;
    }

    #pragma warning disable CA5350 // ETSI EN 319 142-1 mandates SHA-1 only as the VRI dictionary key identifier.
    private static string ComputeVriKey(byte[] pdf, PdfSignatureValidationResult signature) {
        IReadOnlyList<long> byteRange = signature.Signature.ByteRangeValues;
        int start = checked((int)(byteRange[0] + byteRange[1]));
        int end = checked((int)byteRange[2]);
        while (start < end && IsPdfWhitespace(pdf[start])) start++;
        while (end > start && IsPdfWhitespace(pdf[end - 1])) end--;
        if (end - start < 2 || pdf[start] != (byte)'<' || pdf[end - 1] != (byte)'>') {
            throw new NotSupportedException("PAdES VRI keys require a hexadecimal signature /Contents value.");
        }

        var contentsToken = new byte[end - start];
        Buffer.BlockCopy(pdf, start, contentsToken, 0, contentsToken.Length);
        byte[] hash;
#if NET5_0_OR_GREATER
        hash = SHA1.HashData(contentsToken);
        return Convert.ToHexString(hash);
#else
        using (SHA1 sha1 = SHA1.Create()) {
            hash = sha1.ComputeHash(contentsToken);
        }

        const string hex = "0123456789ABCDEF";
        var characters = new char[hash.Length * 2];
        for (int i = 0; i < hash.Length; i++) {
            characters[i * 2] = hex[hash[i] >> 4];
            characters[(i * 2) + 1] = hex[hash[i] & 0x0F];
        }

        return new string(characters);
#endif
    }
    #pragma warning restore CA5350

    private static bool IsPdfWhitespace(byte value) =>
        value == 0 || value == 9 || value == 10 || value == 12 || value == 13 || value == 32;
}
