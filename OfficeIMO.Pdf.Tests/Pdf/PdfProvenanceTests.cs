using System.Text;
using OfficeIMO.Pdf;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfProvenanceTests {
    [Fact]
    public void InspectAndRemoveUseTheExactC2paAssociatedFileProfile() {
        byte[] manifest = CreateManifestStore();
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("PDF provenance"))
            .AttachFile(new PdfEmbeddedFile(
                "keep.txt",
                Encoding.UTF8.GetBytes("keep"),
                "text/plain",
                PdfAssociatedFileRelationship.Supplement))
            .AttachFile(new PdfEmbeddedFile(
                "content-credential.c2pa",
                manifest,
                "application/c2pa",
                PdfAssociatedFileRelationship.C2paManifest))
            .ToBytes();

        OfficeProvenanceReport report = PdfProvenance.Inspect(pdf);
        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(pdf);
        IReadOnlyList<PdfExtractedAttachment> attachments = PdfAttachmentExtractor.ExtractAttachments(result.ToArray());

        OfficeProvenanceEvidence evidence = Assert.Single(report.Evidence);
        Assert.True(evidence.IsStructurallyValid);
        Assert.Equal(OfficeProvenanceAssetFormat.Pdf, report.Format);
        Assert.Empty(result.After.Evidence);
        Assert.All(result.Changes, change => Assert.Equal(0, change.RemovedBytes));
        PdfExtractedAttachment retained = Assert.Single(attachments);
        Assert.Equal("keep.txt", retained.FileName);
        Assert.Equal("keep", Encoding.UTF8.GetString(retained.Bytes));
        Assert.Equal(PdfAssociatedFileRelationship.Supplement, retained.Relationship);
    }

    [Fact]
    public void MalformedC2paAssociatedFileIsPreservedByDefault() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Malformed provenance"))
            .AttachFile(new PdfEmbeddedFile(
                "content-credential.c2pa",
                Encoding.ASCII.GetBytes("not-a-manifest"),
                "application/c2pa",
                PdfAssociatedFileRelationship.C2paManifest))
            .ToBytes();

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(pdf);

        Assert.False(result.WasChanged);
        Assert.Single(result.Before.Evidence);
        Assert.False(result.Before.Evidence[0].IsStructurallyValid);
        Assert.Equal(pdf, result.ToArray());
    }

    [Fact]
    public void CallerCanExplicitlyRemoveMalformedCandidateCarrier() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Malformed provenance"))
            .AttachFile(new PdfEmbeddedFile(
                "content-credential.c2pa",
                Encoding.ASCII.GetBytes("not-a-manifest"),
                "application/c2pa",
                PdfAssociatedFileRelationship.C2paManifest))
            .ToBytes();
        var options = new OfficeProvenanceRemovalOptions { RequireStructurallyValidCarrier = false };

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(pdf, options);

        Assert.True(result.WasChanged);
        Assert.Empty(PdfAttachmentExtractor.ExtractAttachments(result.ToArray()));
    }

    [Fact]
    public void InspectionEnforcesTheSharedAssetLimitBeforePdfParsing() {
        byte[] pdf = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Bounded")).ToBytes();

        Assert.Throws<InvalidDataException>(() => PdfProvenance.Inspect(
            pdf,
            new OfficeProvenanceOptions {
                MaxAssetBytes = pdf.Length - 1,
                MaxManifestBytes = Math.Min(64, pdf.Length - 1)
            }));
    }

    [Fact]
    public void CandidateWithoutAnAssociatedFileReferenceIsNotStructurallyValid() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Detached PDF provenance"))
            .AttachFile(new PdfEmbeddedFile(
                "content-credential.c2pa",
                CreateManifestStore(),
                "application/c2pa",
                PdfAssociatedFileRelationship.C2paManifest))
            .ToBytes();
        byte[] detached = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            catalog.Items.Remove("AF");
            return security.InfoObjectNumber;
        });

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(detached);

        Assert.Single(result.Before.Evidence);
        Assert.False(result.Before.Evidence[0].IsStructurallyValid);
        Assert.False(result.WasChanged);
        Assert.Equal(detached, result.ToArray());
    }

    [Fact]
    public void PageAssociatedFileReferenceDoesNotSatisfyTheDocumentCredentialProfile() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] pageAssociated = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfObject associatedFiles = catalog.Items["AF"];
            catalog.Items.Remove("AF");
            PdfDictionary page = Assert.IsType<PdfDictionary>(objects.Values
                .Select(item => item.Value)
                .First(value => value is PdfDictionary dictionary && dictionary.Get<PdfName>("Type")?.Name == "Page"));
            page.Items["AF"] = associatedFiles;
            return security.InfoObjectNumber;
        });

        OfficeProvenanceReport report = PdfProvenance.Inspect(pageAssociated);

        OfficeProvenanceEvidence evidence = Assert.Single(report.Evidence);
        Assert.False(evidence.IsStructurallyValid);
    }

    [Fact]
    public void ObjectLevelInformationResourceAssociationIsStructurallyValid() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] objectAssociated = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray catalogAssociations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, catalogAssociations, "content-credential.c2pa");
            catalogAssociations.Items.Remove(candidate);
            PdfStream contentStream = objects.Values.Select(item => item.Value).OfType<PdfStream>()
                .First(stream => stream.Dictionary.Get<PdfName>("Type")?.Name != "EmbeddedFile");
            var objectAssociations = new PdfArray();
            objectAssociations.Items.Add(candidate);
            contentStream.Dictionary.Items["AF"] = objectAssociations;
            return security.InfoObjectNumber;
        });

        OfficeProvenanceReport report = PdfProvenance.Inspect(objectAssociated);

        Assert.True(Assert.Single(report.Evidence).IsStructurallyValid);
    }

    [Fact]
    public void UntypedFileSpecificationSelfAssociationIsStructurallyInvalid() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] selfAssociated = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray catalogAssociations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, catalogAssociations, "content-credential.c2pa");
            catalog.Items.Remove("AF");
            PdfDictionary fileSpecification = Assert.IsType<PdfDictionary>(objects[candidate.ObjectNumber].Value);
            fileSpecification.Items.Remove("Type");
            var selfAssociation = new PdfArray();
            selfAssociation.Items.Add(candidate);
            fileSpecification.Items["AF"] = selfAssociation;
            return security.InfoObjectNumber;
        });

        OfficeProvenanceReport report = PdfProvenance.Inspect(selfAssociated);

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
    }

    [Fact]
    public void CatalogReachabilityHandlesDeepIndirectChainsIteratively() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] deeplyLinked = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            int firstObjectNumber = objects.Keys.Max() + 1;
            const int chainLength = 20_000;
            for (int index = 0; index < chainLength; index++) {
                var node = new PdfDictionary();
                if (index + 1 < chainLength) node.Items["Next"] = new PdfReference(firstObjectNumber + index + 1, 0);
                objects[firstObjectNumber + index] = new PdfIndirectObject(firstObjectNumber + index, 0, node);
            }
            catalog.Items["ProvenanceExtension"] = new PdfReference(firstObjectNumber, 0);
            return security.InfoObjectNumber;
        });

        OfficeProvenanceReport report = PdfProvenance.Inspect(deeplyLinked);

        Assert.True(Assert.Single(report.Evidence).IsStructurallyValid);
    }

    [Fact]
    public void CatalogAssociationWithoutNameTreeOrAnnotationIsStructurallyInvalid() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] catalogOnly = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfDictionary names = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, catalog.Items["Names"]));
            names.Items.Remove("EmbeddedFiles");
            return security.InfoObjectNumber;
        });

        OfficeProvenanceReport report = PdfProvenance.Inspect(catalogOnly);

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
    }

    [Fact]
    public void OrphanFileAttachmentAnnotationDoesNotSatisfyTheDocumentProfile() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] orphanReferenced = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray catalogAssociations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, catalogAssociations, "content-credential.c2pa");
            PdfDictionary names = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, catalog.Items["Names"]));
            names.Items.Remove("EmbeddedFiles");
            var orphanAnnotation = new PdfDictionary();
            orphanAnnotation.Items["Type"] = new PdfName("Annot");
            orphanAnnotation.Items["Subtype"] = new PdfName("FileAttachment");
            orphanAnnotation.Items["FS"] = candidate;
            int orphanNumber = objects.Keys.Max() + 1;
            objects[orphanNumber] = new PdfIndirectObject(orphanNumber, 0, orphanAnnotation);
            return orphanNumber;
        });

        OfficeProvenanceReport report = PdfProvenance.Inspect(orphanReferenced);

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
    }

    [Fact]
    public void RemovalPreservesUnrelatedObjectLevelAssociationSites() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] objectAssociated = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray catalogAssociations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference retained = FindFileSpecReference(objects, catalogAssociations, "keep.txt");
            catalogAssociations.Items.Remove(retained);
            PdfStream contentStream = objects.Values.Select(item => item.Value).OfType<PdfStream>()
                .First(stream => stream.Dictionary.Get<PdfName>("Type")?.Name != "EmbeddedFile");
            var objectAssociations = new PdfArray();
            objectAssociations.Items.Add(retained);
            contentStream.Dictionary.Items["AF"] = objectAssociations;
            return security.InfoObjectNumber;
        });

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(objectAssociated);
        var parsed = PdfSyntax.ParseObjects(result.ToArray());
        PdfDictionary rewrittenCatalog = Assert.IsType<PdfDictionary>(PdfSyntax.FindCatalog(parsed.Map, parsed.TrailerRaw));
        Assert.False(rewrittenCatalog.Items.ContainsKey("AF"));
        PdfStream rewrittenContent = parsed.Map.Values.Select(item => item.Value).OfType<PdfStream>()
            .First(stream => stream.Dictionary.Get<PdfName>("Type")?.Name != "EmbeddedFile");
        PdfArray retainedAssociations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(parsed.Map, rewrittenContent.Dictionary.Items["AF"]));
        Assert.Single(retainedAssociations.Items);
        Assert.Equal("keep.txt", GetFileSpecName(parsed.Map, Assert.IsType<PdfReference>(retainedAssociations.Items[0])));
    }

    [Fact]
    public void RemovalPreservesHierarchicalEmbeddedFilesNameTreeWithUpdatedLimits() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] hierarchical = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfDictionary names = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, catalog.Items["Names"]));
            PdfDictionary rootTree = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, names.Items["EmbeddedFiles"]));
            PdfArray flatNames = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, rootTree.Items["Names"]));
            PdfObject[] pairs = flatNames.Items.ToArray();
            rootTree.Items.Remove("Names");
            var kids = new PdfArray();
            int nextNumber = objects.Keys.Max() + 1;
            for (int index = 0; index + 1 < pairs.Length; index += 2) {
                var childNames = new PdfArray();
                childNames.Items.Add(pairs[index]);
                childNames.Items.Add(pairs[index + 1]);
                var limits = new PdfArray();
                limits.Items.Add(pairs[index]);
                limits.Items.Add(pairs[index]);
                var child = new PdfDictionary();
                child.Items["Names"] = childNames;
                child.Items["Limits"] = limits;
                objects[nextNumber] = new PdfIndirectObject(nextNumber, 0, child);
                kids.Items.Add(new PdfReference(nextNumber++, 0));
            }
            rootTree.Items["Kids"] = kids;
            var rootLimits = new PdfArray();
            rootLimits.Items.Add(pairs[0]);
            rootLimits.Items.Add(pairs[pairs.Length - 2]);
            rootTree.Items["Limits"] = rootLimits;
            return security.InfoObjectNumber;
        });

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(hierarchical);
        var parsed = PdfSyntax.ParseObjects(result.ToArray());
        PdfDictionary catalog = Assert.IsType<PdfDictionary>(PdfSyntax.FindCatalog(parsed.Map, parsed.TrailerRaw));
        PdfDictionary names = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(parsed.Map, catalog.Items["Names"]));
        PdfDictionary rootTree = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(parsed.Map, names.Items["EmbeddedFiles"]));

        Assert.False(rootTree.Items.ContainsKey("Names"));
        PdfArray retainedKids = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(parsed.Map, rootTree.Items["Kids"]));
        PdfDictionary retainedChild = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(parsed.Map, Assert.Single(retainedKids.Items)));
        PdfArray retainedNames = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(parsed.Map, retainedChild.Items["Names"]));
        Assert.Equal(2, retainedNames.Items.Count);
        Assert.Equal("keep.txt", Assert.IsType<PdfStringObj>(retainedNames.Items[0]).Value);
        PdfArray limits = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(parsed.Map, rootTree.Items["Limits"]));
        Assert.Equal("keep.txt", Assert.IsType<PdfStringObj>(limits.Items[0]).Value);
        Assert.Equal("keep.txt", Assert.IsType<PdfStringObj>(limits.Items[1]).Value);
    }

    [Fact]
    public void RemovalMapsDuplicateCarrierDescriptorsToDistinctAttachmentIndices() {
        byte[] duplicated = DuplicateCandidateAroundRetainedAttachment(CreatePdfWithCandidateAndRetainedAttachment(), copies: 2);

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(duplicated);
        IReadOnlyList<PdfExtractedAttachment> retained = PdfAttachmentExtractor.ExtractAttachments(result.ToArray());

        Assert.Equal(2, result.Before.Evidence.Count);
        PdfExtractedAttachment attachment = Assert.Single(retained);
        Assert.Equal("keep.txt", attachment.FileName);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void AttachmentRewriteBudgetCountsEveryProjectedDescriptorCopy() {
        byte[] duplicated = DuplicateCandidateAroundRetainedAttachment(CreatePdfWithCandidateAndRetainedAttachment(), copies: 3);

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfAttachmentEditor.Edit(duplicated, _ => { }, maxDecodedAttachmentBytes: 64));

        Assert.Equal(PdfReadLimitKind.AttachmentBytes, exception.Kind);
    }

    [Fact]
    public void InspectionDoesNotDecodeUnrelatedAttachmentsAgainstTheProvenanceBudget() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Bounded provenance"))
            .AttachFile(new PdfEmbeddedFile(
                "unrelated.bin",
                new byte[1024 * 1024],
                "application/octet-stream",
                PdfAssociatedFileRelationship.Supplement))
            .AttachFile(new PdfEmbeddedFile(
                "content-credential.c2pa",
                CreateManifestStore(),
                "application/c2pa",
                PdfAssociatedFileRelationship.C2paManifest))
            .ToBytes();
        var options = new OfficeProvenanceOptions {
            MaxAssetBytes = pdf.Length + 1L,
            MaxManifestBytes = 64,
            MaxExpandedContainerBytes = 64
        };

        OfficeProvenanceReport report = PdfProvenance.Inspect(pdf, options);

        Assert.Single(report.Evidence);
        Assert.True(report.Evidence[0].IsStructurallyValid);
    }

    [Fact]
    public void InspectionAppliesTheManifestLimitBeforeDecodingEachCandidate() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Bounded provenance manifest"))
            .AttachFile(new PdfEmbeddedFile(
                "content-credential.c2pa",
                new byte[4096],
                "application/c2pa",
                PdfAssociatedFileRelationship.C2paManifest))
            .ToBytes();
        var options = new OfficeProvenanceOptions {
            MaxAssetBytes = pdf.Length + 1L,
            MaxManifestBytes = 64,
            MaxExpandedContainerBytes = 1024 * 1024
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => PdfProvenance.Inspect(pdf, options));

        Assert.Equal(PdfReadLimitKind.AttachmentBytes, exception.Kind);
        Assert.Equal(64, exception.Limit);
    }

    [Fact]
    public void RemovalEnforcesExpandedContainerLimitDuringGraphRewrite() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        var options = new OfficeProvenanceRemovalOptions();
        options.Limits.MaxAssetBytes = pdf.LongLength + 1L;
        options.Limits.MaxManifestBytes = 64;
        options.Limits.MaxExpandedContainerBytes = 128;

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => PdfProvenance.Remove(pdf, options));

        Assert.Contains("expanded container limit", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void RemovalPreservesDirectFileSpecificationsInTheEmbeddedFilesNameTree() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] directFileSpecification = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfDictionary names = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, catalog.Items["Names"]));
            PdfDictionary embeddedFiles = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, names.Items["EmbeddedFiles"]));
            PdfArray entries = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, embeddedFiles.Items["Names"]));
            for (int index = 0; index + 1 < entries.Items.Count; index += 2) {
                if (Assert.IsType<PdfStringObj>(entries.Items[index]).Value != "keep.txt") continue;
                PdfReference reference = Assert.IsType<PdfReference>(entries.Items[index + 1]);
                entries.Items[index + 1] = Assert.IsType<PdfDictionary>(objects[reference.ObjectNumber].Value);
            }
            return security.InfoObjectNumber;
        });

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(directFileSpecification);
        var parsed = PdfSyntax.ParseObjects(result.ToArray());
        PdfDictionary catalog = Assert.IsType<PdfDictionary>(PdfSyntax.FindCatalog(parsed.Map, parsed.TrailerRaw));
        PdfDictionary names = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(parsed.Map, catalog.Items["Names"]));
        PdfDictionary embeddedFiles = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(parsed.Map, names.Items["EmbeddedFiles"]));
        PdfArray entries = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(parsed.Map, embeddedFiles.Items["Names"]));

        Assert.Equal(2, entries.Items.Count);
        Assert.Equal("keep.txt", Assert.IsType<PdfStringObj>(entries.Items[0]).Value);
        PdfDictionary retained = Assert.IsType<PdfDictionary>(entries.Items[1]);
        Assert.Equal("keep.txt", retained.Get<PdfStringObj>("UF")?.Value ?? retained.Get<PdfStringObj>("F")?.Value);
    }

    [Fact]
    public void RemovalDropsAnOwnerReferenceToAnEmptiedIndirectAssociatedFilesArray() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] indirectAssociatedFiles = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray catalogAssociations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, catalogAssociations, "content-credential.c2pa");
            var candidateOnly = new PdfArray();
            candidateOnly.Items.Add(candidate);
            int nextObjectNumber = objects.Keys.Max() + 1;
            objects[nextObjectNumber] = new PdfIndirectObject(nextObjectNumber, 0, candidateOnly);
            catalog.Items["AF"] = new PdfReference(nextObjectNumber, 0);
            return security.InfoObjectNumber;
        });

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(indirectAssociatedFiles);
        var parsed = PdfSyntax.ParseObjects(result.ToArray());
        PdfDictionary catalog = Assert.IsType<PdfDictionary>(PdfSyntax.FindCatalog(parsed.Map, parsed.TrailerRaw));

        Assert.False(catalog.Items.ContainsKey("AF"));
        Assert.Empty(result.After.Evidence);
    }

    private static byte[] CreatePdfWithCandidateAndRetainedAttachment() => PdfDocument.Create()
        .Paragraph(paragraph => paragraph.Text("PDF provenance"))
        .AttachFile(new PdfEmbeddedFile(
            "keep.txt",
            Encoding.UTF8.GetBytes("keep"),
            "text/plain",
            PdfAssociatedFileRelationship.Supplement))
        .AttachFile(new PdfEmbeddedFile(
            "content-credential.c2pa",
            CreateManifestStore(),
            "application/c2pa",
            PdfAssociatedFileRelationship.C2paManifest))
        .ToBytes();

    private static byte[] DuplicateCandidateAroundRetainedAttachment(byte[] pdf, int copies) =>
        PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfDictionary names = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, catalog.Items["Names"]));
            PdfDictionary embeddedFiles = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, names.Items["EmbeddedFiles"]));
            PdfArray entries = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, embeddedFiles.Items["Names"]));
            var pairs = new List<(PdfStringObj Name, PdfObject Reference)>();
            for (int index = 0; index + 1 < entries.Items.Count; index += 2) {
                pairs.Add((Assert.IsType<PdfStringObj>(entries.Items[index]), entries.Items[index + 1]));
            }
            (PdfStringObj Name, PdfObject Reference) candidate = pairs.Single(pair => pair.Name.Value.EndsWith(".c2pa", StringComparison.Ordinal));
            (PdfStringObj Name, PdfObject Reference) retained = pairs.Single(pair => pair.Name.Value == "keep.txt");
            entries.Items.Clear();
            for (int index = 0; index < copies; index++) {
                entries.Items.Add(new PdfStringObj(candidate.Name.Value + index, true));
                entries.Items.Add(candidate.Reference);
                if (index == 0) {
                    entries.Items.Add(retained.Name);
                    entries.Items.Add(retained.Reference);
                }
            }
            return security.InfoObjectNumber;
        });

    private static PdfReference FindFileSpecReference(
        Dictionary<int, PdfIndirectObject> objects,
        PdfArray references,
        string fileName) => references.Items.OfType<PdfReference>()
            .Single(reference => GetFileSpecName(objects, reference) == fileName);

    private static string? GetFileSpecName(Dictionary<int, PdfIndirectObject> objects, PdfReference reference) =>
        (PdfObjectLookup.Resolve(objects, reference) as PdfDictionary)?.Get<PdfStringObj>("UF")?.Value ??
        (PdfObjectLookup.Resolve(objects, reference) as PdfDictionary)?.Get<PdfStringObj>("F")?.Value;

    private static byte[] CreateManifestStore() {
        byte[] data = new byte[38];
        WriteBigEndian(data, 0, data.Length);
        Encoding.ASCII.GetBytes("jumb").CopyTo(data, 4);
        WriteBigEndian(data, 8, 30);
        Encoding.ASCII.GetBytes("jumd").CopyTo(data, 12);
        new byte[] { 0x63, 0x32, 0x70, 0x61, 0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71 }.CopyTo(data, 16);
        data[32] = 0x02;
        Encoding.ASCII.GetBytes("c2pa").CopyTo(data, 33);
        return data;
    }

    private static void WriteBigEndian(byte[] data, int offset, int value) {
        data[offset] = (byte)(value >> 24);
        data[offset + 1] = (byte)(value >> 16);
        data[offset + 2] = (byte)(value >> 8);
        data[offset + 3] = (byte)value;
    }
}
