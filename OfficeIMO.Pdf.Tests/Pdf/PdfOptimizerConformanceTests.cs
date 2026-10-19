using OfficeIMO.Pdf;
using Xunit;
#if NET8_0_OR_GREATER
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Security;
#endif

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfOptimizerConformanceTests {
    [Theory]
    [InlineData(PdfOptimizationProfile.Custom)]
    [InlineData(PdfOptimizationProfile.Balanced)]
    [InlineData(PdfOptimizationProfile.MaximumCompression)]
    [InlineData(PdfOptimizationProfile.Web)]
    [InlineData(PdfOptimizationProfile.Archival)]
    public void Optimize_PreservesRawConformanceStreamsAndSourceInfo(PdfOptimizationProfile profile) {
        byte[] source = BuildConformanceFixture();
        var before = PdfSyntax.ParseObjects(source).Map;
        PdfDictionary sourceCatalog = Catalog(before);
        PdfStream sourceMetadata = Resolve<PdfStream>(before, sourceCatalog.Items["Metadata"]);
        PdfExtractedAttachment sourceAttachment = Assert.Single(PdfAttachmentExtractor.ExtractAttachments(source));
        PdfStream sourceEmbedded = Assert.IsType<PdfStream>(before[sourceAttachment.EmbeddedFileObjectNumber].Value);
        Assert.False(sourceMetadata.Dictionary.Items.ContainsKey("Filter"));
        Assert.False(sourceEmbedded.Dictionary.Items.ContainsKey("Filter"));
        Assert.True(sourceEmbedded.Data.Length >= 8192);
        Assert.Contains("urn:officeimo:optimizer:custom", Encoding.UTF8.GetString(sourceMetadata.Data), StringComparison.Ordinal);

        PdfOptimizationActionResult result = PdfOptimizer.Optimize(source, RewriteOptions(profile));

        Assert.False(result.ReturnedOriginal);
        Assert.True(result.PreservationReport.IsPreserved);
        var after = PdfSyntax.ParseObjects(result.Bytes).Map;
        PdfDictionary outputCatalog = Catalog(after);
        PdfStream outputMetadata = Resolve<PdfStream>(after, outputCatalog.Items["Metadata"]);
        AssertStreamUnchanged(sourceMetadata, before, outputMetadata, after);
        Assert.False(outputMetadata.Dictionary.Items.ContainsKey("Filter"));

        PdfExtractedAttachment attachment = Assert.Single(PdfAttachmentExtractor.ExtractAttachments(result.Bytes));
        Assert.Equal(sourceAttachment.Bytes, attachment.Bytes);
        Assert.Equal(sourceAttachment.FileName, attachment.FileName);
        Assert.Equal(sourceAttachment.UnicodeFileName, attachment.UnicodeFileName);
        Assert.Equal("text/plain", attachment.MimeType);
        Assert.Equal(PdfAssociatedFileRelationship.Data, attachment.Relationship);
        PdfStream outputEmbedded = Assert.IsType<PdfStream>(after[attachment.EmbeddedFileObjectNumber].Value);
        AssertStreamUnchanged(sourceEmbedded, before, outputEmbedded, after);
        Assert.False(outputEmbedded.Dictionary.Items.ContainsKey("Filter"));
        Assert.Equal(Fingerprint(before[sourceAttachment.FileSpecObjectNumber].Value, before),
            Fingerprint(after[attachment.FileSpecObjectNumber].Value, after));
        PdfReference associatedFile = Assert.IsType<PdfReference>(Assert.Single(Resolve<PdfArray>(after, outputCatalog.Items["AF"]).Items));
        Assert.Equal(attachment.FileSpecObjectNumber, associatedFile.ObjectNumber);

        PdfArray sourceIntents = Resolve<PdfArray>(before, sourceCatalog.Items["OutputIntents"]);
        PdfArray outputIntents = Resolve<PdfArray>(after, outputCatalog.Items["OutputIntents"]);
        Assert.Equal(2, sourceIntents.Items.Count);
        Assert.Equal(2, outputIntents.Items.Count);
        var profileReferences = new List<int>();
        for (int i = 0; i < sourceIntents.Items.Count; i++) {
            PdfDictionary sourceIntent = Resolve<PdfDictionary>(before, sourceIntents.Items[i]);
            PdfDictionary outputIntent = Resolve<PdfDictionary>(after, outputIntents.Items[i]);
            Assert.Equal(Fingerprint(sourceIntent, before), Fingerprint(outputIntent, after));
            PdfStream sourceIcc = Resolve<PdfStream>(before, sourceIntent.Items["DestOutputProfile"]);
            PdfStream outputIcc = Resolve<PdfStream>(after, outputIntent.Items["DestOutputProfile"]);
            AssertStreamUnchanged(sourceIcc, before, outputIcc, after);
            profileReferences.Add(Assert.IsType<PdfReference>(outputIntent.Items["DestOutputProfile"]).ObjectNumber);
        }
        Assert.Equal(profileReferences[0], profileReferences[1]);
        Assert.Contains(result.Actions, action => action.Kind == "DeduplicateStream");

        int sourceInfo = PdfInspector.Inspect(source).Security.InfoObjectNumber!.Value;
        int outputInfo = PdfInspector.Inspect(result.Bytes).Security.InfoObjectNumber!.Value;
        PdfDictionary info = Assert.IsType<PdfDictionary>(before[sourceInfo].Value);
        Assert.True(info.Items.ContainsKey("PrivateWorkflowData"));
        Assert.Equal(Fingerprint(info, before), Fingerprint(after[outputInfo].Value, after));
    }

    [Theory]
    [InlineData(PdfOptimizationProfile.Custom)]
    [InlineData(PdfOptimizationProfile.Balanced)]
    [InlineData(PdfOptimizationProfile.MaximumCompression)]
    [InlineData(PdfOptimizationProfile.Web)]
    [InlineData(PdfOptimizationProfile.Archival)]
    public void Optimize_PdfA4GroundworkDoesNotCreateInfo(PdfOptimizationProfile profile) {
        byte[] source = PdfDocument.Create(new PdfOptions().ConfigurePdfAGroundwork(PdfComplianceProfile.PdfA4))
            .Meta(title: "XMP-only archival title", author: "OfficeIMO")
            .Paragraph(p => p.Text("PDF/A-4 metadata preservation"))
            .ToBytes();
        Assert.Null(PdfInspector.Inspect(source).Security.InfoObjectNumber);

        PdfOptimizationActionResult result = PdfOptimizer.Optimize(source, RewriteOptions(profile));

        Assert.False(result.ReturnedOriginal);
        PdfDocumentInfo info = PdfInspector.Inspect(result.Bytes);
        Assert.Null(info.Security.InfoObjectNumber);
        Assert.Equal(4, info.XmpMetadata!.PdfAPart);
        var before = PdfSyntax.ParseObjects(source).Map;
        var after = PdfSyntax.ParseObjects(result.Bytes).Map;
        AssertStreamUnchanged(Resolve<PdfStream>(before, Catalog(before).Items["Metadata"]), before,
            Resolve<PdfStream>(after, Catalog(after).Items["Metadata"]), after);
    }

    [Theory]
    [InlineData(PdfOptimizationProfile.MaximumCompression)]
    [InlineData(PdfOptimizationProfile.Custom)]
    public void Optimize_PdfA1RejectsStreamXrefBeforeReplacingOutput(PdfOptimizationProfile profile) {
        byte[] source = BuildPdfA1Groundwork();
        PdfOptimizationOptions options = RewriteOptions(profile);
        if (profile == PdfOptimizationProfile.Custom) {
            options.UseObjectStreams = false;
            options.XrefFormat = PdfOptimizationXrefFormat.XrefStream;
        }
        WithPreservedFiles(source, (input, output) => {
            NotSupportedException exception = Assert.Throws<NotSupportedException>(() => PdfOptimizer.Optimize(input, output, options));
            Assert.Contains("PDF/A-1", exception.Message, StringComparison.Ordinal);
            Assert.Contains("ClassicTable", exception.Message, StringComparison.Ordinal);
        });
    }

    [Theory]
    [InlineData(PdfOptimizationProfile.Custom)]
    [InlineData(PdfOptimizationProfile.Balanced)]
    [InlineData(PdfOptimizationProfile.Web)]
    [InlineData(PdfOptimizationProfile.Archival)]
    public void Optimize_PdfA1KeepsPdf14AndClassicObjects(PdfOptimizationProfile profile) {
        byte[] source = BuildPdfA1Groundwork();

        PdfOptimizationActionResult result = PdfOptimizer.Optimize(source, RewriteOptions(profile));

        Assert.False(result.ReturnedOriginal);
        PdfDocumentInfo info = PdfInspector.Inspect(result.Bytes);
        Assert.Equal("1.4", PdfSyntax.GetHeaderVersion(result.Bytes));
        Assert.Equal(1, info.XmpMetadata!.PdfAPart);
        Assert.False(info.Security.HasObjectStreams);
        Assert.False(info.Security.HasXrefStreams);
        var before = PdfSyntax.ParseObjects(source).Map;
        var after = PdfSyntax.ParseObjects(result.Bytes).Map;
        AssertStreamUnchanged(Resolve<PdfStream>(before, Catalog(before).Items["Metadata"]), before,
            Resolve<PdfStream>(after, Catalog(after).Items["Metadata"]), after);
    }

    [Theory]
    [InlineData(PdfOptimizationProfile.Custom)]
    [InlineData(PdfOptimizationProfile.Balanced)]
    [InlineData(PdfOptimizationProfile.MaximumCompression)]
    [InlineData(PdfOptimizationProfile.Web)]
    [InlineData(PdfOptimizationProfile.Archival)]
    public void Optimize_RefusesFilteredXmpWithoutChangingFiles(PdfOptimizationProfile profile) {
        byte[] original = BuildConformanceFixture();
        var objects = PdfSyntax.ParseObjects(original).Map;
        PdfReference reference = Assert.IsType<PdfReference>(Catalog(objects).Items["Metadata"]);
        PdfStream metadata = Resolve<PdfStream>(objects, reference);
        byte[] compressed = PdfFlateEncoder.Compress(metadata.Data, default);
        metadata.Dictionary.Items["Filter"] = new PdfName("FlateDecode");
        ReplaceStream(objects, reference.ObjectNumber, compressed);
        byte[] source = AssembleFixture(objects, original, PdfFileVersion.Pdf17);

        WithPreservedFiles(source, (input, output) => {
            NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
                PdfOptimizer.Optimize(input, output, RewriteOptions(profile)));
            Assert.Contains("uncompressed XMP", exception.Message, StringComparison.Ordinal);
        });
    }

    [Theory]
    [InlineData(PdfOptimizationProfile.MaximumCompression, false)]
    [InlineData(PdfOptimizationProfile.Custom, false)]
    [InlineData(PdfOptimizationProfile.Custom, true)]
    public void Optimize_RefusesStreamLayoutsWhenXmpExceedsReadLimit(PdfOptimizationProfile profile, bool objectStreams) {
        byte[] original = BuildPdfA1Groundwork();
        var objects = PdfSyntax.ParseObjects(original).Map;
        PdfReference reference = Assert.IsType<PdfReference>(Catalog(objects).Items["Metadata"]);
        PdfStream metadata = Resolve<PdfStream>(objects, reference);
        byte[] padding = Encoding.UTF8.GetBytes(new string(' ', PdfReadDocument.MaxXmpMetadataBytes));
        byte[] packet = new byte[metadata.Data.Length + padding.Length];
        Buffer.BlockCopy(metadata.Data, 0, packet, 0, metadata.Data.Length);
        Buffer.BlockCopy(padding, 0, packet, metadata.Data.Length, padding.Length);
        ReplaceStream(objects, reference.ObjectNumber, packet);
        byte[] source = AssembleFixture(objects, original, PdfFileVersion.Pdf14);
        PdfXmpMetadataInfo xmp = Assert.IsType<PdfXmpMetadataInfo>(PdfInspector.Inspect(source).XmpMetadata);
        Assert.False(xmp.IsWellFormedXml);
        Assert.Null(xmp.PdfAPart);
        PdfOptimizationOptions options = RewriteOptions(profile);
        if (profile == PdfOptimizationProfile.Custom) {
            options.UseObjectStreams = objectStreams;
            options.XrefFormat = objectStreams ? PdfOptimizationXrefFormat.ClassicTable : PdfOptimizationXrefFormat.XrefStream;
        }

        WithPreservedFiles(source, (input, output) => {
            NotSupportedException exception = Assert.Throws<NotSupportedException>(() => PdfOptimizer.Optimize(input, output, options));
            Assert.Contains("readable XMP", exception.Message, StringComparison.Ordinal);
        });
    }

#if NET8_0_OR_GREATER
    [Theory]
    [InlineData(PdfOptimizationProfile.Custom)]
    [InlineData(PdfOptimizationProfile.Balanced)]
    [InlineData(PdfOptimizationProfile.MaximumCompression)]
    [InlineData(PdfOptimizationProfile.Web)]
    [InlineData(PdfOptimizationProfile.Archival)]
    public void Optimize_RejectsCryptographicallySignedPdfWithoutChangingFiles(PdfOptimizationProfile profile) {
        using RSA key = RSA.Create(2048);
        var request = new CertificateRequest("CN=OfficeIMO optimizer regression", key,
            HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, true));
        using X509Certificate2 certificate = request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddDays(1));
        using var signer = new PdfCmsExternalSigner(OfficeSecurityProvider.Default, certificate);
        PdfExternalSignaturePreparation preparation = PdfDocument.Create()
            .Paragraph(p => p.Text("Signed optimizer input"))
            .Security.PrepareExternalSignature(new PdfExternalSignatureOptions { ReservedSignatureContentsBytes = 8192 });
        byte[] signed = preparation.Complete(signer.Sign(new PdfExternalSignatureRequest(preparation))).ToBytes();
        var verification = new CmsVerificationOptions();
        verification.CertificateValidation.ChainEvaluator = (_, _) => true;
        var provider = new PdfCmsSignatureCryptographyProvider(OfficeSecurityProvider.Default, verification);
        PdfSignatureValidationReport validation = PdfSignatureValidator.Validate(signed, provider);
        Assert.True(validation.MathematicalSignaturesVerified);
        Assert.True(validation.DigestVerified);

        WithPreservedFiles(signed, (input, output) => {
            PdfMutationBlockedException exception = Assert.Throws<PdfMutationBlockedException>(() =>
                PdfOptimizer.Optimize(input, output, RewriteOptions(profile)));
            Assert.Equal(PdfMutationOperation.Optimize, exception.Plan.Operation);
        });
    }
#endif

    private static PdfOptimizationOptions RewriteOptions(PdfOptimizationProfile profile) {
        PdfOptimizationOptions options = PdfOptimizationOptions.Create(profile);
        options.KeepOriginalWhenNotSmaller = false;
        return options;
    }

    private static byte[] BuildConformanceFixture() {
        byte[] source = PdfRewritePreservationTestSupport.BuildPreservationProofPdf();
        var objects = PdfSyntax.ParseObjects(source).Map;
        PdfDictionary catalog = Catalog(objects);
        PdfReference metadataReference = Assert.IsType<PdfReference>(catalog.Items["Metadata"]);
        PdfStream metadata = Resolve<PdfStream>(objects, metadataReference);
        string customXml = "<rdf:Description rdf:about=\"\" xmlns:custom=\"urn:officeimo:optimizer:custom\">" +
            "<custom:payload><![CDATA[arbitrary XML <tag> & \u0142\u00f3d\u017a " + new string('X', 4096) +
            "]]></custom:payload><!-- preserve this comment --></rdf:Description>";
        ReplaceStream(objects, metadataReference.ObjectNumber,
            Encoding.UTF8.GetBytes(Encoding.UTF8.GetString(metadata.Data).Replace("</rdf:RDF>", customXml + "</rdf:RDF>")));
        metadata.Dictionary.Items["PrivatePacketTag"] = new PdfStringObj("unchanged packet dictionary");

        PdfExtractedAttachment attachment = Assert.Single(PdfAttachmentExtractor.ExtractAttachments(source));
        byte[] attachmentBytes = Encoding.UTF8.GetBytes(new string('A', 8192) + "\nAttachment payload\n");
        ReplaceStream(objects, attachment.EmbeddedFileObjectNumber, attachmentBytes);
        PdfStream embedded = Assert.IsType<PdfStream>(objects[attachment.EmbeddedFileObjectNumber].Value);
        embedded.Dictionary.Items.Remove("Filter");
        embedded.Dictionary.Items.Remove("DecodeParms");
        if (embedded.Dictionary.Items.TryGetValue("Params", out PdfObject? parameters)) {
            Resolve<PdfDictionary>(objects, parameters).Items["Size"] = new PdfNumber(attachmentBytes.Length);
        }

        PdfArray intents = Resolve<PdfArray>(objects, catalog.Items["OutputIntents"]);
        PdfDictionary intent = Resolve<PdfDictionary>(objects, Assert.Single(intents.Items));
        PdfStream icc = Resolve<PdfStream>(objects, intent.Items["DestOutputProfile"]);
        int duplicateId = objects.Keys.Max() + 1;
        var duplicateDictionary = new PdfDictionary();
        foreach (var entry in icc.Dictionary.Items) duplicateDictionary.Items.Add(entry.Key, entry.Value);
        objects.Add(duplicateId, new PdfIndirectObject(duplicateId, 0, new PdfStream(duplicateDictionary, (byte[])icc.Data.Clone())));
        var secondIntent = new PdfDictionary();
        foreach (var entry in intent.Items) secondIntent.Items.Add(entry.Key, entry.Value);
        secondIntent.Items["DestOutputProfile"] = new PdfReference(duplicateId, 0);
        secondIntent.Items["OutputConditionIdentifier"] = new PdfStringObj("Same ICC, second output condition");
        intents.Items.Add(secondIntent);

        int infoId = PdfInspector.Inspect(source).Security.InfoObjectNumber!.Value;
        PdfDictionary info = Assert.IsType<PdfDictionary>(objects[infoId].Value);
        var custom = new PdfDictionary();
        custom.Items["Stage"] = new PdfName("Approved");
        custom.Items["Revision"] = new PdfNumber(7);
        custom.Items["OpaqueBytes"] = new PdfStringObj(new byte[] { 0, 255, 40, 41, 92, 128 });
        info.Items["PrivateWorkflowData"] = custom;
        info.Items["CreationDate"] = new PdfStringObj("D:20200102123456+02'00'");
        return AssembleFixture(objects, source, PdfFileVersion.Pdf17);
    }

    private static byte[] BuildPdfA1Groundwork() {
        var options = new PdfOptions().ConfigurePdfAGroundwork(PdfComplianceProfile.PdfA2B);
        options.FileVersion = PdfFileVersion.Pdf14;
        byte[] source = PdfDocument.Create(options).Paragraph(p => p.Text("PDF/A-1 groundwork")).ToBytes();
        var objects = PdfSyntax.ParseObjects(source).Map;
        PdfReference reference = Assert.IsType<PdfReference>(Catalog(objects).Items["Metadata"]);
        PdfStream metadata = Resolve<PdfStream>(objects, reference);
        string xml = Encoding.UTF8.GetString(metadata.Data);
        Assert.Contains("<pdfaid:part>2</pdfaid:part>", xml, StringComparison.Ordinal);
        ReplaceStream(objects, reference.ObjectNumber,
            Encoding.UTF8.GetBytes(xml.Replace("<pdfaid:part>2</pdfaid:part>", "<pdfaid:part>1</pdfaid:part>")));
        byte[] result = AssembleFixture(objects, source, PdfFileVersion.Pdf14);
        Assert.Equal(1, PdfInspector.Inspect(result).XmpMetadata!.PdfAPart);
        return result;
    }

    private static void ReplaceStream(Dictionary<int, PdfIndirectObject> objects, int id, byte[] data) {
        PdfIndirectObject original = objects[id];
        PdfDictionary dictionary = Assert.IsType<PdfStream>(original.Value).Dictionary;
        dictionary.Items["Length"] = new PdfNumber(data.Length);
        objects[id] = new PdfIndirectObject(id, original.Generation, new PdfStream(dictionary, data));
    }

    private static byte[] AssembleFixture(Dictionary<int, PdfIndirectObject> objects, byte[] original, PdfFileVersion version) {
        PdfDocumentSecurityInfo security = PdfInspector.Inspect(original).Security;
        int[] ids = objects.Keys.OrderBy(id => id).ToArray();
        var map = ids.Select((id, index) => (id, number: index + 1)).ToDictionary(item => item.id, item => item.number);
        var context = new PdfPageExtractor.SerializationContext(map, 0, new Dictionary<int, Dictionary<string, PdfObject>>(), objects);
        byte[][] serialized = ids.Select(id => PdfPageExtractor.SerializeIndirectObject(map[id], objects[id].Value, context)).ToArray();
        return PdfFileAssembler.Assemble(serialized, map[security.RootObjectNumber!.Value],
            security.InfoObjectNumber.HasValue ? map[security.InfoObjectNumber.Value] : 0, version);
    }

    private static PdfDictionary Catalog(Dictionary<int, PdfIndirectObject> objects) =>
        Assert.Single(objects.Values.Select(item => item.Value).OfType<PdfDictionary>(),
            dictionary => dictionary.Get<PdfName>("Type")?.Name == "Catalog");

    private static T Resolve<T>(Dictionary<int, PdfIndirectObject> objects, PdfObject value) where T : PdfObject =>
        Assert.IsType<T>(PdfObjectLookup.ResolveChain(objects, value));

    private static byte[] Fingerprint(PdfObject value, Dictionary<int, PdfIndirectObject> objects) {
        using var fingerprint = new PdfObjectGraphFingerprint(objects, maximumDepth: 64, maximumNodes: 10000);
        fingerprint.AppendRoot(value);
        return fingerprint.Complete();
    }

    private static void AssertStreamUnchanged(PdfStream expected, Dictionary<int, PdfIndirectObject> before,
        PdfStream actual, Dictionary<int, PdfIndirectObject> after) {
        Assert.Equal(expected.Data, actual.Data);
        Assert.Equal(Fingerprint(expected.Dictionary, before), Fingerprint(actual.Dictionary, after));
    }

    private static void WithPreservedFiles(byte[] source, Action<string, string> action) {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO-optimizer-conformance-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string input = Path.Combine(directory, "input.pdf");
            string output = Path.Combine(directory, "existing-output.pdf");
            byte[] sentinel = Encoding.ASCII.GetBytes("Existing output must survive refused optimization.");
            File.WriteAllBytes(input, source);
            File.WriteAllBytes(output, sentinel);
            action(input, output);
            Assert.Equal(source, File.ReadAllBytes(input));
            Assert.Equal(sentinel, File.ReadAllBytes(output));
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }
}
