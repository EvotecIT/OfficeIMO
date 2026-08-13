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
    public void EmbeddedFilePayloadStreamCannotServeAsAnInformationResource() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] embeddedPayloadAssociated = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray catalogAssociations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, catalogAssociations, "content-credential.c2pa");
            catalog.Items.Remove("AF");
            PdfStream embeddedFile = objects.Values.Select(item => item.Value).OfType<PdfStream>()
                .First(stream => stream.Dictionary.Get<PdfName>("Type")?.Name == "EmbeddedFile");
            var associated = new PdfArray();
            associated.Items.Add(candidate);
            embeddedFile.Dictionary.Items["AF"] = associated;
            return security.InfoObjectNumber;
        });

        OfficeProvenanceReport report = PdfProvenance.Inspect(embeddedPayloadAssociated);

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
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
    public void UntypedFileAttachmentAnnotationDoesNotBecomeAnInformationResource() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] annotationAssociated = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray catalogAssociations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, catalogAssociations, "content-credential.c2pa");
            catalog.Items.Remove("AF");
            var annotationAssociations = new PdfArray();
            annotationAssociations.Items.Add(candidate);
            var annotation = new PdfDictionary();
            annotation.Items["Subtype"] = new PdfName("FileAttachment");
            annotation.Items["FS"] = candidate;
            annotation.Items["AF"] = annotationAssociations;
            int annotationNumber = objects.Keys.Max() + 1;
            objects[annotationNumber] = new PdfIndirectObject(annotationNumber, 0, annotation);
            catalog.Items["ProvenanceAnnotation"] = new PdfReference(annotationNumber, 0);
            return security.InfoObjectNumber;
        });

        OfficeProvenanceReport report = PdfProvenance.Inspect(annotationAssociated);

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
    }

    [Fact]
    public void ActiveUntypedFileAttachmentAnnotationCannotMasqueradeAsTheFileSpecification() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] annotationCarrier = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray catalogAssociations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, catalogAssociations, "content-credential.c2pa");
            PdfDictionary fileSpecification = Assert.IsType<PdfDictionary>(objects[candidate.ObjectNumber].Value);
            fileSpecification.Items.Remove("Type");
            fileSpecification.Items["Subtype"] = new PdfName("FileAttachment");
            fileSpecification.Items["FS"] = candidate;
            PdfDictionary page = Assert.IsType<PdfDictionary>(objects.Values
                .Select(item => item.Value)
                .First(value => value is PdfDictionary dictionary && dictionary.Get<PdfName>("Type")?.Name == "Page"));
            var annotations = new PdfArray();
            annotations.Items.Add(candidate);
            page.Items["Annots"] = annotations;
            return security.InfoObjectNumber;
        });

        OfficeProvenanceReport report = PdfProvenance.Inspect(annotationCarrier);
        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(annotationCarrier);

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void MalformedEmbeddedFilesNameTreeKeyDoesNotValidateCatalogAssociation() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] malformedNameTree = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfDictionary names = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, catalog.Items["Names"]));
            PdfDictionary embeddedFiles = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, names.Items["EmbeddedFiles"]));
            PdfArray entries = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, embeddedFiles.Items["Names"]));
            for (int index = 0; index + 1 < entries.Items.Count; index += 2) {
                if (entries.Items[index + 1] is not PdfReference reference ||
                    GetFileSpecName(objects, reference) != "content-credential.c2pa") continue;
                entries.Items[index] = new PdfName("MalformedNameTreeKey");
            }
            return security.InfoObjectNumber;
        });

        OfficeProvenanceReport report = PdfProvenance.Inspect(malformedNameTree);

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
    public void StaleGenerationFileAttachmentAnnotationDoesNotSatisfyTheDocumentProfile() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] validAnnotation = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray catalogAssociations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, catalogAssociations, "content-credential.c2pa");
            PdfDictionary names = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, catalog.Items["Names"]));
            names.Items.Remove("EmbeddedFiles");
            PdfDictionary page = Assert.IsType<PdfDictionary>(objects.Values.Select(item => item.Value)
                .First(value => value is PdfDictionary dictionary && dictionary.Get<PdfName>("Type")?.Name == "Page"));
            var annotation = new PdfDictionary();
            annotation.Items["Type"] = new PdfName("Annot");
            annotation.Items["Subtype"] = new PdfName("FileAttachment");
            annotation.Items["FS"] = candidate;
            int annotationNumber = objects.Keys.Max() + 1;
            objects[annotationNumber] = new PdfIndirectObject(annotationNumber, 0, annotation);
            var annotations = new PdfArray();
            annotations.Items.Add(new PdfReference(annotationNumber, 0));
            page.Items["Annots"] = annotations;
            return security.InfoObjectNumber;
        });
        string validText = Encoding.ASCII.GetString(validAnnotation);
        System.Text.RegularExpressions.Match activeReference = System.Text.RegularExpressions.Regex.Match(
            validText,
            @"/FS\s+\d+\s+0\s+R",
            System.Text.RegularExpressions.RegexOptions.CultureInvariant);
        Assert.True(activeReference.Success);
        string staleReference = System.Text.RegularExpressions.Regex.Replace(
            activeReference.Value,
            @"\s0\s+R$",
            " 1 R",
            System.Text.RegularExpressions.RegexOptions.CultureInvariant);
        byte[] staleAnnotation = Encoding.ASCII.GetBytes(validText.Remove(activeReference.Index, activeReference.Length)
            .Insert(activeReference.Index, staleReference));

        OfficeProvenanceReport report = PdfProvenance.Inspect(staleAnnotation);

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
    }

    [Fact]
    public void ProvenanceInspectionEnforcesThePerPageAnnotationLimit() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] withAnnotations = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary page = Assert.IsType<PdfDictionary>(objects.Values.Select(item => item.Value)
                .First(value => value is PdfDictionary dictionary && dictionary.Get<PdfName>("Type")?.Name == "Page"));
            var annotations = new PdfArray();
            annotations.Items.Add(new PdfDictionary());
            annotations.Items.Add(new PdfDictionary());
            page.Items["Annots"] = annotations;
            return security.InfoObjectNumber;
        });
        var readOptions = new PdfReadOptions { Limits = new PdfReadLimits { MaxAnnotationsPerPage = 1 } };

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            PdfProvenance.Inspect(withAnnotations, readOptions: readOptions));

        Assert.Contains("annotation", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ActivePageDictionaryCannotMasqueradeAsAFileSpecification() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] contradictory = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray catalogAssociations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, catalogAssociations, "content-credential.c2pa");
            PdfDictionary fileSpecification = Assert.IsType<PdfDictionary>(objects[candidate.ObjectNumber].Value);
            fileSpecification.Items["Type"] = new PdfName("Page");
            return security.InfoObjectNumber;
        });

        OfficeProvenanceReport report = PdfProvenance.Inspect(contradictory);
        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(contradictory);

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void UntypedActivePageTreeNodeCannotMasqueradeAsAFileSpecification() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] contradictory = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray catalogAssociations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, catalogAssociations, "content-credential.c2pa");
            PdfDictionary candidateFileSpec = Assert.IsType<PdfDictionary>(objects[candidate.ObjectNumber].Value);
            PdfReference pagesReference = Assert.IsType<PdfReference>(catalog.Items["Pages"]);
            PdfDictionary pages = Assert.IsType<PdfDictionary>(objects[pagesReference.ObjectNumber].Value);
            pages.Items.Remove("Type");
            pages.Items["EF"] = candidateFileSpec.Items["EF"];
            pages.Items["AFRelationship"] = candidateFileSpec.Items["AFRelationship"];
            pages.Items["F"] = candidateFileSpec.Items["F"];
            pages.Items["UF"] = candidateFileSpec.Items["UF"];
            catalogAssociations.Items.Clear();
            catalogAssociations.Items.Add(pagesReference);
            PdfDictionary names = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, catalog.Items["Names"]));
            PdfDictionary embeddedFiles = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, names.Items["EmbeddedFiles"]));
            PdfArray entries = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, embeddedFiles.Items["Names"]));
            for (int index = 0; index + 1 < entries.Items.Count; index += 2) {
                if (Assert.IsType<PdfStringObj>(entries.Items[index]).Value == "content-credential.c2pa") {
                    entries.Items[index + 1] = pagesReference;
                }
            }
            return security.InfoObjectNumber;
        });

        OfficeProvenanceReport report = PdfProvenance.Inspect(contradictory);
        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(contradictory);

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void FileSpecificationWithUnselectedEmbeddedFileVariantIsPreserved() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] mixedVariants = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray catalogAssociations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, catalogAssociations, "content-credential.c2pa");
            PdfDictionary fileSpec = Assert.IsType<PdfDictionary>(objects[candidate.ObjectNumber].Value);
            PdfDictionary embeddedFiles = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, fileSpec.Items["EF"]));
            int unrelatedObjectNumber = objects.Keys.Max() + 1;
            var unrelatedDictionary = new PdfDictionary();
            unrelatedDictionary.Items["Type"] = new PdfName("EmbeddedFile");
            unrelatedDictionary.Items["Subtype"] = new PdfName("text#2Fplain");
            objects[unrelatedObjectNumber] = new PdfIndirectObject(
                unrelatedObjectNumber,
                0,
                new PdfStream(unrelatedDictionary, Encoding.UTF8.GetBytes("keep-variant")));
            embeddedFiles.Items["F"] = new PdfReference(unrelatedObjectNumber, 0);
            return security.InfoObjectNumber;
        });

        OfficeProvenanceReport report = PdfProvenance.Inspect(mixedVariants);
        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(mixedVariants);

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
        Assert.Equal(mixedVariants, result.ToArray());
    }

    [Fact]
    public void CatalogExtensionFileAttachmentDoesNotSatisfyTheDocumentProfile() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] extensionReferenced = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray catalogAssociations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, catalogAssociations, "content-credential.c2pa");
            PdfDictionary names = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, catalog.Items["Names"]));
            names.Items.Remove("EmbeddedFiles");
            var extensionAnnotation = new PdfDictionary();
            extensionAnnotation.Items["Type"] = new PdfName("Annot");
            extensionAnnotation.Items["Subtype"] = new PdfName("FileAttachment");
            extensionAnnotation.Items["FS"] = candidate;
            catalog.Items["OfficeIMOExtension"] = extensionAnnotation;
            return security.InfoObjectNumber;
        });

        OfficeProvenanceReport report = PdfProvenance.Inspect(extensionReferenced);

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
    public void RemovalPreservesCompleteNameTreePairsBeforeADanglingKey() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] danglingKey = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfDictionary names = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, catalog.Items["Names"]));
            PdfDictionary embeddedFiles = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, names.Items["EmbeddedFiles"]));
            PdfArray entries = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, embeddedFiles.Items["Names"]));
            entries.Items.Add(new PdfStringObj("dangling", true));
            return security.InfoObjectNumber;
        });

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(danglingKey);
        var parsed = PdfSyntax.ParseObjects(result.ToArray());
        PdfDictionary catalog = Assert.IsType<PdfDictionary>(PdfSyntax.FindCatalog(parsed.Map, parsed.TrailerRaw));
        PdfDictionary names = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(parsed.Map, catalog.Items["Names"]));
        PdfDictionary embeddedFiles = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(parsed.Map, names.Items["EmbeddedFiles"]));
        PdfArray entries = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(parsed.Map, embeddedFiles.Items["Names"]));

        Assert.Equal(3, entries.Items.Count);
        Assert.Equal("keep.txt", Assert.IsType<PdfStringObj>(entries.Items[0]).Value);
        Assert.Equal("dangling", Assert.IsType<PdfStringObj>(entries.Items[2]).Value);
    }

    [Fact]
    public void RemovalDropsAnEntireMalformedNameTreePairWhenTheTargetOccupiesTheKeySlot() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] malformed = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray catalogAssociations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, catalogAssociations, "content-credential.c2pa");
            PdfDictionary names = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, catalog.Items["Names"]));
            PdfDictionary embeddedFiles = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, names.Items["EmbeddedFiles"]));
            PdfArray entries = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, embeddedFiles.Items["Names"]));
            entries.Items.Insert(0, PdfNull.Instance);
            entries.Items.Insert(0, candidate);
            return security.InfoObjectNumber;
        });

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(malformed);
        var parsed = PdfSyntax.ParseObjects(result.ToArray());
        PdfDictionary catalog = Assert.IsType<PdfDictionary>(PdfSyntax.FindCatalog(parsed.Map, parsed.TrailerRaw));
        PdfDictionary names = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(parsed.Map, catalog.Items["Names"]));
        PdfDictionary embeddedFiles = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(parsed.Map, names.Items["EmbeddedFiles"]));
        PdfArray entries = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(parsed.Map, embeddedFiles.Items["Names"]));

        Assert.Equal(2, entries.Items.Count);
        Assert.Equal("keep.txt", Assert.IsType<PdfStringObj>(entries.Items[0]).Value);
        Assert.Equal("keep.txt", GetFileSpecName(parsed.Map, Assert.IsType<PdfReference>(entries.Items[1])));
    }

    [Fact]
    public void RemovalDeletesPopupsLinkedToRemovedFileAttachmentAnnotations() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        int attachmentNumber = 0;
        int popupNumber = 0;
        byte[] withPopup = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray catalogAssociations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, catalogAssociations, "content-credential.c2pa");
            PdfDictionary page = Assert.IsType<PdfDictionary>(objects.Values.Select(item => item.Value)
                .First(value => value is PdfDictionary dictionary && dictionary.Get<PdfName>("Type")?.Name == "Page"));
            attachmentNumber = objects.Keys.Max() + 1;
            popupNumber = attachmentNumber + 1;
            var attachment = new PdfDictionary();
            attachment.Items["Type"] = new PdfName("Annot");
            attachment.Items["Subtype"] = new PdfName("FileAttachment");
            attachment.Items["FS"] = candidate;
            attachment.Items["Popup"] = new PdfReference(popupNumber, 0);
            var popup = new PdfDictionary();
            popup.Items["Type"] = new PdfName("Annot");
            popup.Items["Subtype"] = new PdfName("Popup");
            popup.Items["Parent"] = new PdfReference(attachmentNumber, 0);
            objects[attachmentNumber] = new PdfIndirectObject(attachmentNumber, 0, attachment);
            objects[popupNumber] = new PdfIndirectObject(popupNumber, 0, popup);
            var annotations = new PdfArray();
            annotations.Items.Add(new PdfReference(attachmentNumber, 0));
            annotations.Items.Add(new PdfReference(popupNumber, 0));
            page.Items["Annots"] = annotations;
            return security.InfoObjectNumber;
        });

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(withPopup);
        var parsed = PdfSyntax.ParseObjects(result.ToArray());

        Assert.DoesNotContain(attachmentNumber, parsed.Map.Keys);
        Assert.DoesNotContain(popupNumber, parsed.Map.Keys);
        PdfDictionary page = Assert.IsType<PdfDictionary>(parsed.Map.Values.Select(item => item.Value)
            .First(value => value is PdfDictionary dictionary && dictionary.Get<PdfName>("Type")?.Name == "Page"));
        Assert.False(page.Items.TryGetValue("Annots", out PdfObject? annotationsValue) &&
            PdfObjectLookup.Resolve(parsed.Map, annotationsValue) is PdfArray annotations && annotations.Items.Count != 0);
    }

    [Fact]
    public void RemovalDeletesFileAttachmentAnnotationsFromUntypedPageLeaves() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        int annotationNumber = 0;
        byte[] untypedPage = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray catalogAssociations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, catalogAssociations, "content-credential.c2pa");
            PdfDictionary page = Assert.IsType<PdfDictionary>(objects.Values.Select(item => item.Value)
                .First(value => value is PdfDictionary dictionary && dictionary.Get<PdfName>("Type")?.Name == "Page"));
            page.Items.Remove("Type");
            annotationNumber = objects.Keys.Max() + 1;
            var attachment = new PdfDictionary();
            attachment.Items["Type"] = new PdfName("Annot");
            attachment.Items["Subtype"] = new PdfName("FileAttachment");
            attachment.Items["FS"] = candidate;
            objects[annotationNumber] = new PdfIndirectObject(annotationNumber, 0, attachment);
            var annotations = new PdfArray();
            annotations.Items.Add(new PdfReference(annotationNumber, 0));
            page.Items["Annots"] = annotations;
            return security.InfoObjectNumber;
        });

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(untypedPage);
        var parsed = PdfSyntax.ParseObjects(result.ToArray());

        Assert.True(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.DoesNotContain(annotationNumber, parsed.Map.Keys);
        PdfDictionary page = Assert.IsType<PdfDictionary>(parsed.Map.Values.Select(item => item.Value)
            .First(value => value is PdfDictionary dictionary && dictionary.Items.ContainsKey("MediaBox")));
        Assert.False(page.Items.TryGetValue("Annots", out PdfObject? annotationsValue) &&
            PdfObjectLookup.Resolve(parsed.Map, annotationsValue) is PdfArray annotations && annotations.Items.Count != 0);
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
            MaxManifestBytes = 256,
            MaxExpandedContainerBytes = 256
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
    public void InspectionAppliesCarrierLimitBeforeDecodingTheNextCandidate() {
        byte[] duplicated = DuplicateCandidateAroundRetainedAttachment(CreatePdfWithCandidateAndRetainedAttachment(), copies: 2);
        var options = new OfficeProvenanceOptions {
            MaxAssetBytes = duplicated.LongLength + 1L,
            MaxManifestBytes = 256,
            MaxExpandedContainerBytes = 1024 * 1024,
            MaxCarriers = 1
        };

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => PdfProvenance.Inspect(duplicated, options));

        Assert.Contains("carrier limit", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void RemovalEnforcesExpandedContainerLimitDuringGraphRewrite() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        var options = new OfficeProvenanceRemovalOptions();
        options.Limits.MaxAssetBytes = pdf.LongLength + 1L;
        options.Limits.MaxManifestBytes = 256;
        options.Limits.MaxExpandedContainerBytes = 192;

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => PdfProvenance.Remove(pdf, options));

        Assert.Contains("expanded container limit", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void GraphRewritePreflightsLargeStreamBodiesBeforeSerialization() {
        var stream = new PdfStream(new PdfDictionary(), new byte[8 * 1024 * 1024]);
        var context = new PdfPageExtractor.SerializationContext(
            new Dictionary<int, int>(),
            0,
            new Dictionary<int, Dictionary<string, PdfObject>>());

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            PdfPageExtractor.EnsureSerializedObjectWithinLimit(stream, context, 1024 * 1024));

        Assert.Contains("expanded container limit", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void GraphRewritePreflightUsesExactSerializedSizesForSmallNumbers() {
        var values = new PdfArray();
        for (int index = 0; index < 4096; index++) values.Items.Add(new PdfNumber(0));
        var context = new PdfPageExtractor.SerializationContext(
            new Dictionary<int, int>(),
            0,
            new Dictionary<int, Dictionary<string, PdfObject>>());
        byte[] serialized = PdfPageExtractor.SerializeObject(values, context);

        PdfPageExtractor.EnsureSerializedObjectWithinLimit(values, context, serialized.LongLength);
        PdfPageExtractor.EnsureSerializedIndirectObjectWithinLimit(
            values, context, 1, serialized.LongLength + 15L);

        Assert.Throws<InvalidDataException>(() =>
            PdfPageExtractor.EnsureSerializedObjectWithinLimit(values, context, serialized.LongLength - 1L));
    }

    [Fact]
    public void UntypedEmbeddedFileParameterDictionariesAreNotInformationResourceAssociationSites() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] parametersAssociation = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray catalogAssociations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, catalogAssociations, "content-credential.c2pa");
            PdfDictionary fileSpecification = Assert.IsType<PdfDictionary>(objects[candidate.ObjectNumber].Value);
            PdfDictionary embeddedFiles = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, fileSpecification.Items["EF"]));
            PdfReference embeddedFileReference = Assert.IsType<PdfReference>(
                embeddedFiles.Items.TryGetValue("UF", out PdfObject? unicodeFile) ? unicodeFile : embeddedFiles.Items["F"]);
            PdfStream embeddedFile = Assert.IsType<PdfStream>(objects[embeddedFileReference.ObjectNumber].Value);
            embeddedFile.Dictionary.Items.Remove("Type");
            var parameters = new PdfDictionary();
            var associations = new PdfArray();
            associations.Items.Add(candidate);
            parameters.Items["AF"] = associations;
            int parameterObjectNumber = objects.Keys.Max() + 1;
            objects[parameterObjectNumber] = new PdfIndirectObject(parameterObjectNumber, 0, parameters);
            embeddedFile.Dictionary.Items["Params"] = new PdfReference(parameterObjectNumber, 0);
            catalog.Items.Remove("AF");
            return security.InfoObjectNumber;
        });

        OfficeProvenanceReport report = PdfProvenance.Inspect(parametersAssociation);
        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(parametersAssociation);

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UntypedEmbeddedFileGraphRolesAreNotInformationResourceAssociationSites(bool useStreamDictionary) {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] graphAssociation = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray catalogAssociations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, catalogAssociations, "content-credential.c2pa");
            PdfDictionary fileSpecification = Assert.IsType<PdfDictionary>(objects[candidate.ObjectNumber].Value);
            PdfDictionary embeddedFiles = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, fileSpecification.Items["EF"]));
            PdfReference embeddedFileReference = Assert.IsType<PdfReference>(
                embeddedFiles.Items.TryGetValue("UF", out PdfObject? unicodeFile) ? unicodeFile : embeddedFiles.Items["F"]);
            PdfStream embeddedFile = Assert.IsType<PdfStream>(objects[embeddedFileReference.ObjectNumber].Value);
            embeddedFile.Dictionary.Items.Remove("Type");
            var associations = new PdfArray();
            associations.Items.Add(candidate);
            (useStreamDictionary ? embeddedFile.Dictionary : embeddedFiles).Items["AF"] = associations;
            catalog.Items.Remove("AF");
            return security.InfoObjectNumber;
        });

        OfficeProvenanceReport report = PdfProvenance.Inspect(graphAssociation);
        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(graphAssociation);

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void PageResourceDictionariesAreNotInformationResourceAssociationSites() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] resourcesAssociation = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray catalogAssociations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, catalogAssociations, "content-credential.c2pa");
            catalog.Items.Remove("AF");
            PdfDictionary page = Assert.IsType<PdfDictionary>(objects.Values
                .Select(item => item.Value)
                .First(value => value is PdfDictionary dictionary && dictionary.Get<PdfName>("Type")?.Name == "Page"));
            var resources = new PdfDictionary();
            var associations = new PdfArray();
            associations.Items.Add(candidate);
            resources.Items["AF"] = associations;
            int resourceObjectNumber = objects.Keys.Max() + 1;
            objects[resourceObjectNumber] = new PdfIndirectObject(resourceObjectNumber, 0, resources);
            page.Items["Resources"] = new PdfReference(resourceObjectNumber, 0);
            return security.InfoObjectNumber;
        });

        OfficeProvenanceReport report = PdfProvenance.Inspect(resourcesAssociation);
        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(resourcesAssociation);

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void ActiveResourceDictionaryCannotMasqueradeAsAnUntypedFileSpecification() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] resourceFileSpecification = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray catalogAssociations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, catalogAssociations, "content-credential.c2pa");
            PdfDictionary fileSpecification = Assert.IsType<PdfDictionary>(objects[candidate.ObjectNumber].Value);
            fileSpecification.Items.Remove("Type");
            PdfDictionary page = Assert.IsType<PdfDictionary>(objects.Values
                .Select(item => item.Value)
                .First(value => value is PdfDictionary dictionary && dictionary.Get<PdfName>("Type")?.Name == "Page"));
            page.Items["Resources"] = candidate;
            return security.InfoObjectNumber;
        });

        OfficeProvenanceReport report = PdfProvenance.Inspect(resourceFileSpecification);
        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(resourceFileSpecification);

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ActiveAcroFormFieldsCannotMasqueradeAsUntypedFileSpecifications(bool nestedField) {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] fieldFileSpecification = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray catalogAssociations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, catalogAssociations, "content-credential.c2pa");
            PdfDictionary fileSpecification = Assert.IsType<PdfDictionary>(objects[candidate.ObjectNumber].Value);
            fileSpecification.Items.Remove("Type");
            var fields = new PdfArray();
            if (nestedField) {
                var parent = new PdfDictionary();
                var kids = new PdfArray();
                kids.Items.Add(candidate);
                parent.Items["Kids"] = kids;
                int parentObjectNumber = objects.Keys.Max() + 1;
                objects[parentObjectNumber] = new PdfIndirectObject(parentObjectNumber, 0, parent);
                fields.Items.Add(new PdfReference(parentObjectNumber, 0));
            } else {
                fields.Items.Add(candidate);
            }
            var acroForm = new PdfDictionary();
            acroForm.Items["Fields"] = fields;
            catalog.Items["AcroForm"] = acroForm;
            return security.InfoObjectNumber;
        });

        OfficeProvenanceReport report = PdfProvenance.Inspect(fieldFileSpecification);
        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(fieldFileSpecification);

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void ActiveOutlineItemsCannotMasqueradeAsUntypedFileSpecifications() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] outlineFileSpecification = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray catalogAssociations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, catalogAssociations, "content-credential.c2pa");
            PdfDictionary fileSpecification = Assert.IsType<PdfDictionary>(objects[candidate.ObjectNumber].Value);
            fileSpecification.Items.Remove("Type");
            var outlines = new PdfDictionary();
            outlines.Items["First"] = candidate;
            outlines.Items["Last"] = candidate;
            int outlinesObjectNumber = objects.Keys.Max() + 1;
            objects[outlinesObjectNumber] = new PdfIndirectObject(outlinesObjectNumber, 0, outlines);
            fileSpecification.Items["Parent"] = new PdfReference(outlinesObjectNumber, 0);
            catalog.Items["Outlines"] = new PdfReference(outlinesObjectNumber, 0);
            return security.InfoObjectNumber;
        });

        OfficeProvenanceReport report = PdfProvenance.Inspect(outlineFileSpecification);
        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(outlineFileSpecification);

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void ActiveTrailerInfoDictionaryCannotMasqueradeAsAnUntypedFileSpecification() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] infoFileSpecification = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray catalogAssociations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, catalogAssociations, "content-credential.c2pa");
            PdfDictionary fileSpecification = Assert.IsType<PdfDictionary>(objects[candidate.ObjectNumber].Value);
            fileSpecification.Items.Remove("Type");
            return candidate.ObjectNumber;
        });

        OfficeProvenanceReport report = PdfProvenance.Inspect(infoFileSpecification);
        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(infoFileSpecification);

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
        Assert.Equal(infoFileSpecification, result.ToArray());
    }

    [Fact]
    public void ResourceDictionaryDiscoveryHandlesDeepIndirectChainsIteratively() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] deeplyLinked = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary page = Assert.IsType<PdfDictionary>(objects.Values
                .Select(item => item.Value)
                .First(value => value is PdfDictionary dictionary && dictionary.Get<PdfName>("Type")?.Name == "Page"));
            int firstObjectNumber = objects.Keys.Max() + 1;
            const int chainLength = 10_000;
            for (int index = 0; index < chainLength; index++) {
                var dictionary = new PdfDictionary();
                if (index + 1 < chainLength) dictionary.Items["Next"] = new PdfReference(firstObjectNumber + index + 1, 0);
                objects[firstObjectNumber + index] = new PdfIndirectObject(firstObjectNumber + index, 0, dictionary);
            }
            page.Items["Resources"] = new PdfReference(firstObjectNumber, 0);
            return security.InfoObjectNumber;
        });

        OfficeProvenanceReport report = PdfProvenance.Inspect(
            deeplyLinked,
            new OfficeProvenanceOptions { MaxContainerEntries = 20_000 },
            new PdfReadOptions { Limits = new PdfReadLimits { MaxIndirectObjects = 20_000 } });

        Assert.True(Assert.Single(report.Evidence).IsStructurallyValid);
    }

    [Fact]
    public void CatalogNameTreeDiscoveryHandlesDeepIndirectChainsIteratively() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] deeplyLinked = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfDictionary names = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, catalog.Items["Names"]));
            int firstObjectNumber = objects.Keys.Max() + 1;
            const int chainLength = 10_000;
            for (int index = 0; index < chainLength; index++) {
                var dictionary = new PdfDictionary();
                if (index + 1 < chainLength) {
                    var kids = new PdfArray();
                    kids.Items.Add(new PdfReference(firstObjectNumber + index + 1, 0));
                    dictionary.Items["Kids"] = kids;
                }
                objects[firstObjectNumber + index] = new PdfIndirectObject(firstObjectNumber + index, 0, dictionary);
            }
            names.Items["Custom"] = new PdfReference(firstObjectNumber, 0);
            return security.InfoObjectNumber;
        });

        OfficeProvenanceReport report = PdfProvenance.Inspect(
            deeplyLinked,
            new OfficeProvenanceOptions { MaxContainerEntries = 50_000 },
            new PdfReadOptions { Limits = new PdfReadLimits { MaxIndirectObjects = 20_000 } });

        Assert.True(Assert.Single(report.Evidence).IsStructurallyValid);
    }

    [Fact]
    public void EmbeddedFilesNameTreeDictionariesAreNotInformationResourceAssociationSites() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] treeAssociation = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray catalogAssociations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, catalogAssociations, "content-credential.c2pa");
            catalog.Items.Remove("AF");
            PdfDictionary names = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, catalog.Items["Names"]));
            PdfDictionary embeddedFiles = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, names.Items["EmbeddedFiles"]));
            var associations = new PdfArray();
            associations.Items.Add(candidate);
            embeddedFiles.Items["AF"] = associations;
            return security.InfoObjectNumber;
        });

        OfficeProvenanceReport report = PdfProvenance.Inspect(treeAssociation);
        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(treeAssociation);

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void ProvenanceContainerEntryLimitCapsPdfStructuralParsing() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        var options = new OfficeProvenanceOptions { MaxContainerEntries = 1 };
        var readOptions = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxIndirectObjects = 500_000 }
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfProvenance.Inspect(pdf, options, readOptions));

        Assert.Equal(PdfReadLimitKind.IndirectObjects, exception.Kind);
        Assert.Equal(1, exception.Limit);
    }

    [Fact]
    public void ProvenanceContainerEntryLimitCapsMalformedNameTreeEntries() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] oversizedNameTree = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfDictionary names = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, catalog.Items["Names"]));
            PdfDictionary embeddedFiles = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, names.Items["EmbeddedFiles"]));
            PdfArray entries = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, embeddedFiles.Items["Names"]));
            for (int index = 0; index < 256; index++) {
                entries.Items.Add(new PdfStringObj("invalid-" + index.ToString(System.Globalization.CultureInfo.InvariantCulture), true));
                entries.Items.Add(PdfNull.Instance);
            }
            return security.InfoObjectNumber;
        });
        var options = new OfficeProvenanceOptions { MaxContainerEntries = 128 };

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            PdfProvenance.Inspect(oversizedNameTree, options));

        Assert.Contains("container entry limit", exception.Message, StringComparison.OrdinalIgnoreCase);
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
        byte[] storeDescription = CreateBox("jumd", Join(C2paUuid("c2pa"), new byte[] { 0x02 }, Encoding.ASCII.GetBytes("c2pa\0")));
        byte[] manifestDescription = CreateBox("jumd", Join(C2paUuid("c2ma"), new byte[] { 0x02 }, Encoding.ASCII.GetBytes("m\0")));
        byte[] claimDescription = CreateBox("jumd", Join(C2paUuid("c2cl"), new byte[] { 0x02 }, Encoding.ASCII.GetBytes("c2pa.claim\0")));
        byte[] claim = CreateBox("jumb", Join(claimDescription, CreateBox("cbor", new byte[] { 0xA0 })));
        byte[] signatureDescription = CreateBox("jumd", Join(C2paUuid("c2cs"), new byte[] { 0x02 }, Encoding.ASCII.GetBytes("c2pa.signature\0")));
        byte[] signature = CreateBox("jumb", Join(signatureDescription, CreateBox("cbor", new byte[] { 0xA0 })));
        return CreateBox("jumb", Join(storeDescription, CreateBox("jumb", Join(manifestDescription, claim, signature))));
    }

    private static byte[] C2paUuid(string code) => Join(
        Encoding.ASCII.GetBytes(code),
        new byte[] { 0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71 });

    private static byte[] CreateBox(string type, byte[] payload) {
        byte[] box = new byte[payload.Length + 8];
        WriteBigEndian(box, 0, box.Length);
        Encoding.ASCII.GetBytes(type).CopyTo(box, 4);
        Buffer.BlockCopy(payload, 0, box, 8, payload.Length);
        return box;
    }

    private static byte[] Join(params byte[][] arrays) {
        byte[] output = new byte[arrays.Sum(item => item.Length)];
        int offset = 0;
        foreach (byte[] item in arrays) {
            Buffer.BlockCopy(item, 0, output, offset, item.Length);
            offset += item.Length;
        }
        return output;
    }

    private static void WriteBigEndian(byte[] data, int offset, int value) {
        data[offset] = (byte)(value >> 24);
        data[offset + 1] = (byte)(value >> 16);
        data[offset + 2] = (byte)(value >> 8);
        data[offset + 3] = (byte)value;
    }
}
