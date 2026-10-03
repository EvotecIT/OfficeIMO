using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfStructureAssociatedFileTests {
    private const string FirstMath = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>x</mi></math>";
    private const string SecondMath = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>y</mi></math>";

    [Theory]
    [InlineData(PdfObjectSerializationMode.Buffered)]
    [InlineData(PdfObjectSerializationMode.ForwardOnly)]
    public void Ua1RejectsSemanticAttachmentsWithoutChangingProfileOrWritingOutput(PdfObjectSerializationMode serialization) {
        var options = new PdfOptions { ObjectSerializationMode = serialization };
        options.ConfigurePdfUaGroundwork(PdfComplianceProfile.PdfUa1);
        options.RequireCompliance(PdfComplianceProfile.PdfUa1);
        var document = PdfDocument.Create(options).Canvas(canvas => canvas.Structure(
            PdfCanvasStructureRole.Formula, nested => nested.Text("x", 20, 20, 40, 20),
            Supplement("x.xml", FirstMath, "x")));
        using var destination = new MemoryStream();
        InvalidOperationException error = Assert.Throws<InvalidOperationException>(() => document.Save(destination));
        Assert.Contains("Structure-associated files require PDF 2.0", error.Message);
        Assert.Equal(0, destination.Length);
        Assert.Equal(PdfComplianceProfile.PdfUa1, options.ComplianceProfile);
        Assert.Equal(PdfFileVersion.Pdf17, options.FileVersion);
    }

    [Theory]
    [InlineData(PdfObjectSerializationMode.Buffered)]
    [InlineData(PdfObjectSerializationMode.ForwardOnly)]
    public void FormulasHaveTheirOwnRecoverableSupplementAndVersionFloor(PdfObjectSerializationMode serialization) {
        var options = new PdfOptions {
            TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers,
            FileVersion = PdfFileVersion.Pdf17,
            ObjectSerializationMode = serialization
        };
        var document = PdfDocument.Create(options)
            .Canvas(canvas => canvas
                .Structure(PdfCanvasStructureRole.Formula, nested => nested.Text("x", 20, 20, 40, 20), Supplement("x.xml", FirstMath, "First formula"))
                .Structure(PdfCanvasStructureRole.Formula, nested => nested.Text("y", 20, 60, 40, 20), Supplement("y.xml", SecondMath, "Second formula")));
        using var stream = new MemoryStream();
        document.Save(stream);
        byte[] bytes = stream.ToArray();
        Assert.StartsWith("%PDF-2.0", Encoding.ASCII.GetString(bytes, 0, 8));
        Assert.Equal(PdfFileVersion.Pdf17, options.FileVersion);

        var reopened = PdfDocument.Load(bytes);
        PdfRawDocumentView raw = reopened.Resources.RawStructure();
        PdfRawObjectView[] formulas = raw.Objects.Where(IsFormula).ToArray();
        Assert.Equal(2, formulas.Length);
        Assert.Equal("x.xml", FileName(raw, formulas[0]));
        Assert.Equal("y.xml", FileName(raw, formulas[1]));
        Assert.Equal("First formula", formulas[0].Value.Entries["Alt"].Text);
        PdfRawObjectView catalog = raw.GetObject(raw.CatalogObjectNumber!.Value)!;
        Assert.False(catalog.Value.Entries.ContainsKey("AF"));
        PdfExtractedAttachment[] files = reopened.Attachments.Extract().ToArray();
        Assert.Equal(2, files.Length);
        Assert.Equal(FirstMath, Encoding.UTF8.GetString(files.Single(file => file.FileName == "x.xml").Bytes));
        Assert.Equal(SecondMath, Encoding.UTF8.GetString(files.Single(file => file.FileName == "y.xml").Bytes));
        Assert.All(files, file => {
            Assert.Equal("application/mathml+xml", file.MimeType);
            Assert.Equal(PdfAssociatedFileRelationship.Supplement, file.Relationship);
        });
        Assert.Throws<PdfReadLimitException>(() => reopened.Attachments.Extract(reopened.Inspect().Attachments[0], 1));
    }

    [Fact]
    public void IdenticalSupplementsReusePayloadWithoutMergingFormulas() {
        var source = Supplement("same.xml", FirstMath, "First");
        byte[] bytes = PdfDocument.Create(new PdfOptions { TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers })
            .Canvas(canvas => canvas
                .Structure(PdfCanvasStructureRole.Formula, nested => nested.Text("first", 20, 20, 80, 20), source)
                .Structure(PdfCanvasStructureRole.Formula, nested => nested.Text("second", 20, 60, 80, 20), source))
            .ToBytes();
        PdfRawDocumentView raw = PdfReadDocument.Open(bytes).RawStructure();
        PdfRawObjectView[] formulas = raw.Objects.Where(IsFormula).ToArray();
        Assert.Equal(2, formulas.Length);
        Assert.Equal(Assert.Single(formulas[0].Value.Entries["AF"].Items).ReferenceObjectNumber,
            Assert.Single(formulas[1].Value.Entries["AF"].Items).ReferenceObjectNumber);
        Assert.Single(PdfAttachmentExtractor.ExtractAttachments(bytes));
    }

    [Fact]
    public void AssociationSnapshotsIsolateSourceAndOptionsMutations() {
        byte[] data = Encoding.UTF8.GetBytes(FirstMath);
        var file = new PdfEmbeddedFile("source.xml", data, "application/mathml+xml", PdfAssociatedFileRelationship.Supplement);
        var structure = new PdfCanvasStructureOptions().AddAssociatedFile(file);
        data[0] = 0;
        file.FileName = "changed.xml";
        structure.AssociatedFiles[0].MimeType = "text/plain";
        var document = PdfDocument.Create(new PdfOptions { TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers })
            .Canvas(canvas => canvas.Structure(PdfCanvasStructureRole.Formula, nested => nested.Text("x", 20, 20, 40, 20), structure));
        structure.AddAssociatedFile(new PdfEmbeddedFile("later.xml", Encoding.UTF8.GetBytes(SecondMath), "application/mathml+xml"));
        PdfExtractedAttachment recovered = Assert.Single(PdfAttachmentExtractor.ExtractAttachments(document.ToBytes()));
        Assert.Equal("source.xml", recovered.FileName);
        Assert.Equal("application/mathml+xml", recovered.MimeType);
        Assert.Equal(FirstMath, Encoding.UTF8.GetString(recovered.Bytes));
    }

    [Fact]
    public void ConflictingNamesFailInsteadOfRecoveringTheWrongFormula() {
        var document = PdfDocument.Create(new PdfOptions { TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers })
            .Canvas(canvas => canvas
                .Structure(PdfCanvasStructureRole.Formula, nested => nested.Text("x", 20, 20, 40, 20), Supplement("same.xml", FirstMath, "First"))
                .Structure(PdfCanvasStructureRole.Formula, nested => nested.Text("y", 20, 60, 40, 20), Supplement("same.xml", SecondMath, "Second")));
        Assert.Throws<InvalidOperationException>(() => document.ToBytes());
        var catalogConflict = PdfDocument.Create(new PdfOptions { TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers })
            .AttachFile("same.xml", Encoding.UTF8.GetBytes(FirstMath), "application/mathml+xml")
            .Canvas(canvas => canvas.Structure(PdfCanvasStructureRole.Formula, nested => nested.Text("y", 20, 20, 40, 20), Supplement("same.xml", SecondMath, "Second")));
        Assert.Throws<InvalidOperationException>(() => catalogConflict.ToBytes());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ConflictingNamesDoNotWriteToForwardOnlyDestination(bool catalogCollision) {
        var document = PdfDocument.Create(new PdfOptions {
            TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers,
            ObjectSerializationMode = PdfObjectSerializationMode.ForwardOnly,
            FileVersion = PdfFileVersion.Pdf17
        });
        if (catalogCollision) document.AttachFile("same.xml", Encoding.UTF8.GetBytes(FirstMath), "application/mathml+xml");
        else document.Canvas(canvas => canvas.Structure(PdfCanvasStructureRole.Formula,
            nested => nested.Text("x", 20, 20, 40, 20), Supplement("same.xml", FirstMath, "First")));
        document.Canvas(canvas => canvas.Structure(PdfCanvasStructureRole.Formula,
            nested => nested.Text("y", 20, 20, 40, 20), Supplement("same.xml", SecondMath, "Second")));
        using var destination = new MemoryStream();
        Assert.Throws<InvalidOperationException>(() => document.Save(destination));
        Assert.Equal(0, destination.Length);
    }

    [Theory]
    [InlineData(PdfComplianceProfile.PdfUa1)]
    [InlineData(PdfComplianceProfile.PdfUa2)]
    public void FormulaAccessibilityEvidenceRequiresAnAlternateDescription(PdfComplianceProfile profile) {
        foreach (string? description in new string?[] { null, "x squared" }) {
            var options = new PdfOptions { TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers };
            var document = PdfDocument.Create(options).Canvas(canvas => canvas.Structure(
                PdfCanvasStructureRole.Formula, nested => nested.Text("x", 20, 20, 40, 20),
                new PdfCanvasStructureOptions { AlternativeText = description }));
            var requirement = Assert.Single(document.AssessCompliance(profile).Requirements,
                item => item.Id == "generated-formula-alternative-text");
            Assert.Equal(string.IsNullOrWhiteSpace(description) ? PdfComplianceRequirementStatus.Missing
                : PdfComplianceRequirementStatus.Satisfied, requirement.Status);
        }
    }

    [Fact]
    public void UndatedSupplementOmitsOptionalParametersRatherThanInventingDate() {
        byte[] bytes = PdfDocument.Create(new PdfOptions { TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers })
            .Canvas(canvas => canvas.Structure(PdfCanvasStructureRole.Formula, nested => nested.Text("x", 20, 20, 40, 20), Supplement("x.xml", FirstMath, "First")))
            .ToBytes();
        PdfRawDocumentView raw = PdfReadDocument.Open(bytes).RawStructure();
        PdfRawObjectView stream = Assert.Single(raw.Objects, item => item.Value.Entries.TryGetValue("Type", out var type) && type.Text == "EmbeddedFile");
        Assert.False(stream.Value.Entries.ContainsKey("Params"));
        Assert.Equal(FirstMath, Encoding.UTF8.GetString(Assert.Single(PdfAttachmentExtractor.ExtractAttachments(bytes)).Bytes));
        var metadata = new PdfCanvasStructureOptions().AddAssociatedFile(new PdfEmbeddedFile("dated.xml", Encoding.UTF8.GetBytes(FirstMath), "application/mathml+xml", PdfAssociatedFileRelationship.Supplement, modificationDate: new DateTimeOffset(2026, 1, 1, 0, 0, 0, TimeSpan.Zero)));
        bytes = PdfDocument.Create(new PdfOptions { TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers })
            .Canvas(canvas => canvas.Structure(PdfCanvasStructureRole.Formula, nested => nested.Text("x", 20, 20, 40, 20), metadata)).ToBytes();
        raw = PdfReadDocument.Open(bytes).RawStructure();
        stream = Assert.Single(raw.Objects, item => item.Value.Entries.TryGetValue("Type", out var type) && type.Text == "EmbeddedFile");
        Assert.True(stream.Value.Entries["Params"].Entries.ContainsKey("ModDate"));
    }

    [Fact]
    public void AssociationRequiresMimeAndRejectsDuplicateLocalNames() {
        var options = new PdfCanvasStructureOptions();
        Assert.Throws<ArgumentException>(() => options.AddAssociatedFile(new PdfEmbeddedFile("x.xml", Encoding.UTF8.GetBytes(FirstMath))));
        options.AddAssociatedFile(new PdfEmbeddedFile("x.xml", Encoding.UTF8.GetBytes(FirstMath), "application/mathml+xml"));
        Assert.Throws<ArgumentException>(() => options.AddAssociatedFile(new PdfEmbeddedFile("x.xml", Encoding.UTF8.GetBytes(SecondMath), "application/mathml+xml")));
    }

    [Fact]
    public void SupplementsAcrossPagesReusePayloadAndDoNotBecomeCatalogAssociations() {
        var supplement = Supplement("shared.xml", FirstMath, "Repeated formula");
        byte[] bytes = PdfDocument.Create(document => document.Content(content => content
            .Canvas(canvas => canvas.Structure(PdfCanvasStructureRole.Formula, nested => nested.Text("x", 20, 20, 40, 20), supplement))
            .PageBreak()
            .Canvas(canvas => canvas.Structure(PdfCanvasStructureRole.Formula, nested => nested.Text("x", 20, 20, 40, 20), supplement))),
            new PdfOptions().EnableTaggedPdfCatalogMarkers()).ToBytes();
        PdfRawDocumentView raw = PdfReadDocument.Open(bytes).RawStructure();
        PdfRawObjectView[] formulas = raw.Objects.Where(IsFormula).ToArray();
        Assert.Equal(2, PdfInspector.Inspect(bytes).PageCount);
        Assert.Equal(2, formulas.Length);
        Assert.Equal(Assert.Single(formulas[0].Value.Entries["AF"].Items).ReferenceObjectNumber,
            Assert.Single(formulas[1].Value.Entries["AF"].Items).ReferenceObjectNumber);
        Assert.Single(PdfAttachmentExtractor.ExtractAttachments(bytes));
    }

    [Fact]
    public void DisabledTaggingDoesNotEmitOrphanedStructureAttachments() {
        byte[] bytes = PdfDocument.Create(document => document.Content(content => content
            .Canvas(canvas => canvas.Structure(PdfCanvasStructureRole.Formula, nested => nested.Text("x", 20, 20, 40, 20), Supplement("x.xml", FirstMath, "First")))),
            new PdfOptions { TaggedStructureMode = PdfTaggedStructureMode.None }).ToBytes();
        Assert.Empty(PdfAttachmentExtractor.ExtractAttachments(bytes));
        Assert.DoesNotContain(PdfReadDocument.Open(bytes).RawStructure().Objects, IsFormula);
        Assert.StartsWith("%PDF-1.4", Encoding.ASCII.GetString(bytes, 0, 8));
    }

    [Fact]
    public void MixedCatalogAndStructureFilesUseValidPdf20Metadata() {
        byte[] bytes = PdfDocument.Create(new PdfOptions().EnableTaggedPdfCatalogMarkers())
            .AttachFile("catalog.xml", Encoding.UTF8.GetBytes("<source/>"), "application/xml")
            .Canvas(canvas => canvas.Structure(PdfCanvasStructureRole.Formula, nested => nested.Text("x", 20, 20, 40, 20), Supplement("x.xml", FirstMath, "First"))).ToBytes();
        PdfRawDocumentView raw = PdfReadDocument.Open(bytes).RawStructure();
        PdfRawObjectView[] streams = raw.Objects.Where(item => item.Value.Entries.TryGetValue("Type", out var type) && type.Text == "EmbeddedFile").ToArray();
        Assert.Equal(2, streams.Length);
        Assert.All(streams, stream => {
            Assert.True(stream.Value.Entries.ContainsKey("Subtype"));
            Assert.False(stream.Value.Entries.ContainsKey("Params"));
        });
        Assert.Equal(2, PdfAttachmentExtractor.ExtractAttachments(bytes).Count);
    }

    [Fact]
    public void Pdf20CatalogAssociationsRequireMimeWithoutChangingLegacyOutput() {
        var legacy = PdfDocument.Create().AttachFile("catalog.xml", Encoding.UTF8.GetBytes("<source/>"));
        Assert.Single(PdfAttachmentExtractor.ExtractAttachments(legacy.ToBytes()));
        var mixed = PdfDocument.Create(new PdfOptions().EnableTaggedPdfCatalogMarkers())
            .AttachFile("catalog.xml", Encoding.UTF8.GetBytes("<source/>"))
            .Canvas(canvas => canvas.Structure(PdfCanvasStructureRole.Formula, nested => nested.Text("x", 20, 20, 40, 20), Supplement("x.xml", FirstMath, "First")));
        Assert.Throws<InvalidOperationException>(() => mixed.ToBytes());
    }

    [Fact]
    public void RequiredPdfA3SeesStructureFileDatesBeforeWriting() {
        var document = ArchivalFormula(PdfComplianceProfile.PdfA3B, modificationDate: null);
        PdfComplianceRequirement date = Assert.Single(document.AssessCompliance(PdfComplianceProfile.PdfA3B).Requirements, requirement => requirement.Id == "pdfa-embedded-file-modification-dates");
        Assert.Equal(PdfComplianceRequirementStatus.Missing, date.Status);
        using var destination = new MemoryStream();
        var error = Assert.Throws<InvalidOperationException>(() => document.Save(destination));
        Assert.Contains("pdfa-embedded-file-modification-dates", error.Message);
        Assert.Equal(0, destination.Length);
        Assert.Single(PdfAttachmentExtractor.ExtractAttachments(ArchivalFormula(PdfComplianceProfile.PdfA3B, new DateTimeOffset(2026, 1, 1, 0, 0, 0, TimeSpan.Zero)).ToBytes()));
    }

    [Fact]
    public void RequiredPdfA2SeesEffectiveVersionAndPdfXSeesStructureAttachments() {
        var document = ArchivalFormula(PdfComplianceProfile.PdfA2B, new DateTimeOffset(2026, 1, 1, 0, 0, 0, TimeSpan.Zero));
        PdfComplianceRequirement version = Assert.Single(document.AssessCompliance(PdfComplianceProfile.PdfA2B).Requirements, requirement => requirement.Id == "pdf-file-version");
        Assert.Equal(PdfComplianceRequirementStatus.Missing, version.Status);
        using var destination = new MemoryStream();
        var error = Assert.Throws<InvalidOperationException>(() => document.Save(destination));
        Assert.Contains("pdf-file-version", error.Message);
        Assert.Equal(0, destination.Length);
        PdfComplianceRequirement attachments = Assert.Single(document.AssessCompliance(PdfComplianceProfile.PdfX4).Requirements, requirement => requirement.Id == "pdfx-no-embedded-files");
        Assert.Equal(PdfComplianceRequirementStatus.Missing, attachments.Status);
    }

    private static PdfDocument ArchivalFormula(PdfComplianceProfile profile, DateTimeOffset? modificationDate) {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);
        var options = new PdfOptions { IncludeStandardFontToUnicodeMaps = true }
            .ConfigurePdfAGroundwork(profile, "en-US").RequireCompliance(profile)
            .EnableTaggedPdfCatalogMarkers()
            .EmbedStandardFont(PdfStandardFont.Helvetica, File.ReadAllBytes(fontPath!), "OfficeIMO Source Serif");
        var structure = new PdfCanvasStructureOptions().AddAssociatedFile(new PdfEmbeddedFile("x.xml", Encoding.UTF8.GetBytes(FirstMath), "application/mathml+xml", PdfAssociatedFileRelationship.Supplement, modificationDate: modificationDate));
        return PdfDocument.Create(options).Canvas(canvas => canvas.Structure(PdfCanvasStructureRole.Formula, nested => nested.Text("x", 20, 20, 40, 20), structure));
    }

    private static PdfCanvasStructureOptions Supplement(string fileName, string source, string alternativeText) =>
        new PdfCanvasStructureOptions { AlternativeText = alternativeText }.AddAssociatedFile(
            new PdfEmbeddedFile(fileName, Encoding.UTF8.GetBytes(source), "application/mathml+xml", PdfAssociatedFileRelationship.Supplement));

    private static bool IsFormula(PdfRawObjectView item) => item.Value.Entries.TryGetValue("S", out var role) && role.Text == "Formula";

    private static string? FileName(PdfRawDocumentView raw, PdfRawObjectView formula) =>
        raw.GetObject(Assert.Single(formula.Value.Entries["AF"].Items).ReferenceObjectNumber!.Value)!.Value.Entries["UF"].Text;
}
