using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfLogicalDocumentTests {
    [Fact]
    public void Load_BuildsLogicalPagesWithTextTablesAndImages() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Meta(title: "Logical sample", author: "OfficeIMO")
            .H1("Logical Heading")
            .Paragraph(p => p.Text("Logical readback marker."))
            .Bullets(new[] { "Detected logical bullet" })
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "A-100", "Alpha", "2" },
                new[] { "B-200", "Beta", "14" }
            }, style: new PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 70, 170, 60 },
                HeaderRowCount = 1,
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .Image(CreateMinimalRgbPng(), 18, 18)
            .ToBytes();

        PdfLogicalDocument logical = PdfLogicalDocument.Load(pdf, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        });

        PdfLogicalPage page = Assert.Single(logical.Pages);
        Assert.Equal("Logical sample", logical.Metadata.Title);
        Assert.False(logical.HasCatalogActions);
        Assert.Equal(0, logical.CatalogActionCount);
        Assert.Empty(logical.CatalogActions);
        Assert.Empty(logical.CatalogActionNames);
        Assert.Empty(logical.CatalogActionTypes);
        Assert.Empty(logical.CatalogActionSources);
        Assert.Empty(logical.CatalogActionsByActionType);
        Assert.Empty(logical.CatalogActionsBySource);
        Assert.Empty(logical.GetCatalogActionsByActionType("JavaScript"));
        Assert.Empty(logical.GetCatalogActionsBySource("OpenAction"));
        Assert.False(logical.HasAttachments);
        Assert.Equal(0, logical.AttachmentCount);
        Assert.Empty(logical.Attachments);
        Assert.Empty(logical.AttachmentNames);
        Assert.Empty(logical.AttachmentFileNames);
        Assert.Empty(logical.AttachmentSources);
        Assert.Empty(logical.GetAttachmentsByFileName("note.txt"));
        Assert.Empty(logical.GetAttachmentsBySource("AF"));
        Assert.Empty(logical.GetAttachmentsByRelationship(PdfAssociatedFileRelationship.Data));
        Assert.False(logical.HasReadableOutputIntents);
        Assert.Equal(0, logical.OutputIntentCount);
        Assert.Empty(logical.OutputIntents);
        Assert.Empty(logical.OutputIntentSubtypes);
        Assert.Empty(logical.OutputConditionIdentifiers);
        Assert.Empty(logical.GetOutputIntentsBySubtype("GTS_PDFA1"));
        Assert.Empty(logical.GetOutputIntentsByOutputConditionIdentifier("sRGB"));
        Assert.False(logical.HasReadableXmpMetadata);
        Assert.Null(logical.XmpMetadata);
        Assert.False(logical.HasReadableTaggedContent);
        Assert.Null(logical.TaggedContent);
        Assert.False(logical.HasReadableOptionalContent);
        Assert.False(logical.HasOptionalContentGroups);
        Assert.Equal(0, logical.OptionalContentGroupCount);
        Assert.Null(logical.OptionalContent);
        Assert.Empty(logical.OptionalContentGroups);
        Assert.Empty(logical.OptionalContentGroupNames);
        Assert.Empty(logical.GetOptionalContentGroupsByName("Layer 1"));
        Assert.False(logical.HasPageActions);
        Assert.Equal(0, logical.PageActionCount);
        Assert.Empty(logical.PageActions);
        Assert.Empty(logical.PageActionTypes);
        Assert.Empty(logical.PageActionTriggerNames);
        Assert.Empty(logical.PageActionPaths);
        Assert.Empty(logical.PageActionsByActionType);
        Assert.Empty(logical.PageActionsByTriggerName);
        Assert.Empty(logical.PageActionsByActionPath);
        Assert.Empty(logical.PageActionsByPageNumber);
        Assert.Empty(logical.GetPageActionsByActionType("JavaScript"));
        Assert.Empty(logical.GetPageActionsByTriggerName("O"));
        Assert.Empty(logical.GetPageActionsByActionPath("O.Next"));
        Assert.Empty(logical.GetPageActions(1));
        Assert.Throws<ArgumentOutOfRangeException>(() => logical.GetPageActions(0));
        Assert.True(logical.HasSourcePage(1));
        Assert.Same(page, Assert.Single(logical.PagesBySourcePageNumber[1]));
        Assert.Same(page, Assert.Single(logical.GetPages(1)));
        Assert.Empty(logical.GetPages(2));
        Assert.Throws<ArgumentOutOfRangeException>(() => logical.HasSourcePage(0));
        Assert.Throws<ArgumentOutOfRangeException>(() => logical.GetPages(0));
        PdfLogicalHeading heading = Assert.Single(page.Headings);
        Assert.Equal("Logical Heading", heading.Text);
        Assert.Equal(1, heading.Level);
        Assert.Equal(PdfLogicalElementKind.Heading, heading.Line.Kind);
        Assert.Same(heading, Assert.Single(logical.Headings));
        Assert.Contains(page.TextBlocks, block => Normalize(block.Text).Contains("Logicalreadbackmarker", StringComparison.Ordinal));
        Assert.Contains(logical.TextBlocks, block => Normalize(block.Text).Contains("Logicalreadbackmarker", StringComparison.Ordinal));
        Assert.Contains(page.TextBlocks, block =>
            block.Kind == PdfLogicalElementKind.ListItem &&
            Normalize(block.Text).Contains("Detectedlogicalbullet", StringComparison.Ordinal));
        PdfLogicalListItem listItem = Assert.Single(page.ListItems);
        Assert.Equal("Detected logical bullet", listItem.Text);
        Assert.Equal(1, listItem.Level);
        Assert.NotEmpty(listItem.Marker);
        Assert.Equal(PdfLogicalElementKind.ListItem, listItem.Line.Kind);
        Assert.Same(listItem, Assert.Single(logical.ListItems));
        Assert.Contains(page.Paragraphs, paragraph => Normalize(paragraph.Text).Contains("Logicalreadbackmarker", StringComparison.Ordinal));
        Assert.Contains(logical.Paragraphs, paragraph => Normalize(paragraph.Text).Contains("Logicalreadbackmarker", StringComparison.Ordinal));
        Assert.DoesNotContain(page.Paragraphs, paragraph => Normalize(paragraph.Text).Contains("A-100", StringComparison.Ordinal));

        PdfLogicalTable table = Assert.Single(page.Tables, item => item.Rows.Count >= 3 && item.Columns.Count >= 3);
        Assert.Same(table, Assert.Single(logical.Tables, item => item.Rows.Count >= 3 && item.Columns.Count >= 3));
        Assert.Contains(table.Rows, row => row.Count >= 3 &&
            Normalize(row[0]) == "A-100" &&
            Normalize(row[1]) == "Alpha" &&
            Normalize(row[2]) == "2");
        Assert.Contains(table.Cells, cell =>
            cell.PageNumber == 1 &&
            cell.RowIndex == 1 &&
            cell.ColumnIndex == 0 &&
            Normalize(cell.Text) == "A-100" &&
            cell.Column is not null &&
            cell.Column.From < cell.Column.To);
        Assert.Contains(table.Cells, cell =>
            cell.RowIndex == 2 &&
            cell.ColumnIndex == 2 &&
            Normalize(cell.Text) == "14");

        PdfLogicalImage image = Assert.Single(page.Images);
        Assert.Equal(1, image.PageNumber);
        Assert.Equal(1, image.Width);
        Assert.Equal(1, image.Height);
        Assert.Equal("image/png", image.MimeType);
        PdfImagePlacement placement = Assert.Single(image.Placements);
        Assert.True(image.HasPlacements);
        Assert.Equal(1, placement.PageNumber);
        Assert.Equal(image.ResourceName, placement.ResourceName);
        Assert.True(placement.Width > 0);
        Assert.True(placement.Height > 0);
        Assert.True(placement.IsAxisAligned);
        Assert.Same(image, Assert.Single(logical.Images));

        Assert.True(logical.HasElementKind(PdfLogicalElementKind.Table));
        Assert.True(logical.HasElementKind(PdfLogicalElementKind.Image));
        Assert.True(page.HasElementKind(PdfLogicalElementKind.Heading));
        Assert.True(page.HasElementKind(PdfLogicalElementKind.Image));
        Assert.Same(heading.Line, Assert.Single(page.GetElements(PdfLogicalElementKind.Heading)));
        Assert.Same(table, Assert.Single(logical.GetElements(PdfLogicalElementKind.Table)));
        Assert.Same(image, Assert.Single(logical.ElementsByKind[PdfLogicalElementKind.Image]));
        Assert.Equal(page.Elements, logical.ElementsByPageNumber[1]);
        Assert.Equal(page.Elements, logical.GetElements(1));
        Assert.Empty(logical.GetElements(PdfLogicalElementKind.LinkAnnotation));
        Assert.Empty(page.GetElements(PdfLogicalElementKind.LinkAnnotation));
        Assert.Empty(logical.GetElements(2));
        Assert.Throws<ArgumentOutOfRangeException>(() => logical.GetElements(0));
        Assert.Contains(logical.Elements, element => element.Kind == PdfLogicalElementKind.Table);
        Assert.Contains(logical.Elements, element => element.Kind == PdfLogicalElementKind.Image);
    }

    [Fact]
    public void Load_ReadsOutputIntentProfileMetadata() {
        byte[] pdf = PdfDocument.Create(new PdfOptions().SetSrgbOutputIntent())
            .Paragraph(p => p.Text("Logical output intent readback."))
            .PageBreak()
            .Paragraph(p => p.Text("Second page."))
            .ToBytes();

        PdfLogicalDocument logical = PdfLogicalDocument.Load(pdf);

        Assert.True(logical.HasReadableOutputIntents);
        Assert.Equal(1, logical.OutputIntentCount);
        Assert.Equal(new[] { "GTS_PDFA1" }, logical.OutputIntentSubtypes);
        Assert.Equal(new[] { PdfIccProfiles.SrgbIec6196621OutputConditionIdentifier }, logical.OutputConditionIdentifiers);
        PdfOutputIntentInfo outputIntent = Assert.Single(logical.OutputIntents);
        Assert.Equal("GTS_PDFA1", outputIntent.Subtype);
        Assert.Equal(PdfIccProfiles.SrgbIec6196621OutputConditionIdentifier, outputIntent.OutputConditionIdentifier);
        Assert.True(outputIntent.HasDestinationOutputProfile);
        Assert.Equal(3, outputIntent.DestinationOutputProfileColorComponents);
        Assert.True(outputIntent.DestinationOutputProfileSizeBytes > 128);
        Assert.Equal(outputIntent.DestinationOutputProfileSizeBytes, outputIntent.DestinationOutputProfileDeclaredSizeBytes);
        Assert.Equal("RGB ", outputIntent.DestinationOutputProfileColorSpace);
        Assert.True(outputIntent.DestinationOutputProfileHasIccSignature);
        Assert.Same(outputIntent, Assert.Single(logical.GetOutputIntentsBySubtype("GTS_PDFA1")));
        Assert.Same(outputIntent, Assert.Single(logical.GetOutputIntentsByOutputConditionIdentifier(PdfIccProfiles.SrgbIec6196621OutputConditionIdentifier)));
        Assert.Empty(logical.GetOutputIntentsBySubtype("GTS_PDFX"));
        Assert.Empty(logical.GetOutputIntentsByOutputConditionIdentifier("Office profile"));

        PdfLogicalDocument pageRange = PdfLogicalDocument.LoadPageRanges(pdf, new PdfPageRange(1, 1));
        Assert.False(pageRange.HasReadableOutputIntents);
        Assert.Empty(pageRange.OutputIntents);
    }

    [Fact]
    public void Load_ReadsGeneratedXmpMetadataFields() {
        byte[] pdf = PdfDocument.Create(new PdfOptions()
                .SetPdfAIdentification(3, "B")
                .SetPdfUaIdentification()
                .SetElectronicInvoiceMetadata("EN 16931"))
            .Meta(title: "Logical XMP readback", author: "OfficeIMO", subject: "Logical metadata", keywords: "delta, epsilon")
            .Paragraph(p => p.Text("Logical generated XMP readback."))
            .PageBreak()
            .Paragraph(p => p.Text("Second page."))
            .ToBytes();

        PdfLogicalDocument logical = PdfLogicalDocument.Load(pdf);

        Assert.True(logical.HasReadableXmpMetadata);
        PdfXmpMetadataInfo xmp = Assert.IsType<PdfXmpMetadataInfo>(logical.XmpMetadata);
        Assert.True(xmp.IsWellFormedXml);
        Assert.Equal("Logical XMP readback", xmp.Title);
        Assert.Equal("OfficeIMO", xmp.Creator);
        Assert.Equal("Logical metadata", xmp.Description);
        Assert.Equal(new[] { "delta", "epsilon" }, xmp.Subjects);
        Assert.Equal("OfficeIMO.Pdf", xmp.Producer);
        Assert.Equal(3, xmp.PdfAPart);
        Assert.Equal("B", xmp.PdfAConformance);
        Assert.Equal(1, xmp.PdfUaPart);
        Assert.Equal("INVOICE", xmp.ElectronicInvoiceDocumentType);
        Assert.Equal("factur-x.xml", xmp.ElectronicInvoiceDocumentFileName);
        Assert.Equal("1.0", xmp.ElectronicInvoiceVersion);
        Assert.Equal("EN 16931", xmp.ElectronicInvoiceConformanceLevel);

        PdfLogicalDocument pageRange = PdfLogicalDocument.LoadPageRanges(pdf, new PdfPageRange(1, 1));
        Assert.False(pageRange.HasReadableXmpMetadata);
        Assert.Null(pageRange.XmpMetadata);
    }

    [Fact]
    public void Load_ReadsGeneratedTaggedContentMetadata() {
        byte[] pdf = PdfDocument.Create()
            .TaggedPdfCatalogMarkers()
            .Language("en-US")
            .H1("Logical tagged heading")
            .Paragraph(p => p.Text("Logical tagged paragraph."))
            .PageBreak()
            .Paragraph(p => p.Text("Second page."))
            .ToBytes();

        PdfLogicalDocument logical = PdfLogicalDocument.Load(pdf);

        Assert.True(logical.HasReadableTaggedContent);
        PdfTaggedContentInfo tagged = Assert.IsType<PdfTaggedContentInfo>(logical.TaggedContent);
        Assert.True(tagged.Marked);
        Assert.NotNull(tagged.StructTreeRootObjectNumber);
        Assert.NotNull(tagged.ParentTreeObjectNumber);
        Assert.True(tagged.ParentTreeNextKey > 0);
        Assert.NotEmpty(tagged.RootElementObjectNumbers);
        Assert.True(tagged.ParentTreeEntryCount > 0);
        Assert.True(tagged.StructureElementCount >= 4);
        Assert.Contains("Document", tagged.StructureTypes);
        Assert.Contains("H1", tagged.StructureTypes);
        Assert.Contains("P", tagged.StructureTypes);
        Assert.Contains(tagged.StructureElements, element => element.StructureType == "Document" && element.Language == "en-US");
        Assert.Contains(tagged.StructureElements, element => element.StructureType == "P" && element.MarkedContentReferenceCount > 0);

        PdfLogicalDocument pageRange = PdfLogicalDocument.LoadPageRanges(pdf, new PdfPageRange(1, 1));
        Assert.False(pageRange.HasReadableTaggedContent);
        Assert.Null(pageRange.TaggedContent);
    }

    [Fact]
    public void Load_ReadsCatalogActionsWithoutScriptPayload() {
        PdfLogicalDocument logical = PdfLogicalDocument.Load(BuildCatalogJavaScriptActionPdf());

        Assert.True(logical.HasCatalogActions);
        Assert.Equal(3, logical.CatalogActionCount);
        Assert.Equal(new[] { "Open", "OpenAction", "AA.WC" }, logical.CatalogActionNames);
        Assert.Equal(new[] { "JavaScript", "Launch" }, logical.CatalogActionTypes);
        Assert.Equal(new[] { "Names/JavaScript", "OpenAction", "AA" }, logical.CatalogActionSources);

        PdfCatalogAction nameTreeAction = Assert.Single(logical.GetCatalogActionsBySource("Names/JavaScript"));
        Assert.Equal("Open", nameTreeAction.Name);
        Assert.Equal("JavaScript", nameTreeAction.ActionType);
        Assert.Null(nameTreeAction.TriggerName);

        PdfCatalogAction openAction = Assert.Single(logical.GetCatalogActionsBySource("OpenAction"));
        Assert.Equal("OpenAction", openAction.Name);
        Assert.Equal("JavaScript", openAction.ActionType);

        PdfCatalogAction additionalAction = Assert.Single(logical.GetCatalogActionsByActionType("Launch"));
        Assert.Equal("AA.WC", additionalAction.Name);
        Assert.Equal("AA", additionalAction.Source);
        Assert.Equal("WC", additionalAction.TriggerName);
        Assert.Equal(2, logical.GetCatalogActionsByActionType("JavaScript").Count);
        Assert.Empty(logical.GetCatalogActionsBySource("Missing"));
    }

    [Fact]
    public void Load_ReadsReusedCatalogNextActionsForEachBranch() {
        PdfLogicalDocument logical = PdfLogicalDocument.Load(BuildCatalogJavaScriptActionWithSharedNextActionPdf());

        Assert.True(logical.HasCatalogActions);
        Assert.Equal(3, logical.CatalogActionCount);
        Assert.Equal(new[] { "OpenAction", "OpenAction.Next.0", "OpenAction.Next.1" }, logical.CatalogActionNames);
        Assert.Equal(2, logical.GetCatalogActionsByActionType("Launch").Count);
    }

    [Fact]
    public void Load_ReadsPageActionsWithoutScriptPayload() {
        PdfLogicalDocument logical = PdfLogicalDocument.Load(BuildPageAdditionalActionsPdf());

        Assert.True(logical.HasPageActions);
        Assert.Equal(2, logical.PageActionCount);
        Assert.Equal(new[] { "JavaScript", "Launch" }, logical.PageActionTypes);
        Assert.Equal(new[] { "O", "C" }, logical.PageActionTriggerNames);
        Assert.Equal(new[] { "O", "C" }, logical.PageActionPaths);

        PdfLogicalPage page = Assert.Single(logical.Pages);
        Assert.True(page.HasPageActions);
        Assert.Equal(2, page.PageActionCount);

        PdfPageAction openAction = Assert.Single(logical.GetPageActionsByTriggerName("O"));
        Assert.Equal(1, openAction.PageNumber);
        Assert.Equal("JavaScript", openAction.ActionType);
        Assert.Same(openAction, Assert.Single(logical.GetPageActionsByActionType("JavaScript")));
        Assert.Same(openAction, Assert.Single(logical.GetPageActionsByActionPath("O")));
        Assert.Equal(2, logical.GetPageActions(1).Count);
        Assert.Empty(logical.GetPageActions(2));
        Assert.Empty(logical.GetPageActionsByActionType("GoTo"));
        Assert.Empty(logical.GetPageActionsByTriggerName("D"));
        Assert.Empty(logical.GetPageActionsByActionPath("O.Next"));
    }

    [Fact]
    public void Load_ReadsReusedIndirectPageActionsForEachTrigger() {
        PdfLogicalDocument logical = PdfLogicalDocument.Load(BuildPageAdditionalActionsWithSharedIndirectActionPdf());

        Assert.True(logical.HasPageActions);
        Assert.Equal(2, logical.PageActionCount);
        Assert.Equal(new[] { "O", "C" }, logical.PageActionTriggerNames);
        Assert.Equal(new[] { "O", "C" }, logical.PageActionPaths);

        PdfLogicalPage page = Assert.Single(logical.Pages);
        Assert.Equal(2, page.PageActionCount);
        Assert.Equal("JavaScript", page.PageActions[0].ActionType);
        Assert.Equal("JavaScript", page.PageActions[1].ActionType);
        Assert.Equal("O", page.PageActions[0].TriggerName);
        Assert.Equal("C", page.PageActions[1].TriggerName);
        Assert.Equal(2, logical.GetPageActionsByActionType("JavaScript").Count);
    }

    [Fact]
    public void Load_ReadsPageNextActionsWithoutScriptPayload() {
        PdfLogicalDocument logical = PdfLogicalDocument.Load(BuildPageChainedActionsPdf());

        Assert.True(logical.HasPageActions);
        Assert.Equal(3, logical.PageActionCount);
        Assert.Equal(new[] { "JavaScript", "Launch", "RichMedia" }, logical.PageActionTypes);
        Assert.Equal(new[] { "O" }, logical.PageActionTriggerNames);
        Assert.Equal(new[] { "O", "O.Next.0", "O.Next.1" }, logical.PageActionPaths);

        PdfLogicalPage page = Assert.Single(logical.Pages);
        Assert.True(page.HasPageActions);
        Assert.Equal(3, page.PageActionCount);
        Assert.False(page.PageActions[0].IsChainedAction);
        Assert.True(page.PageActions[1].IsChainedAction);

        PdfPageAction richMediaAction = Assert.Single(logical.GetPageActionsByActionPath("O.Next.1"));
        Assert.Equal("RichMedia", richMediaAction.ActionType);
        Assert.Same(richMediaAction, Assert.Single(logical.GetPageActionsByActionType("RichMedia")));
        Assert.Equal(3, logical.GetPageActionsByTriggerName("O").Count);
        Assert.Empty(logical.GetPageActionsByActionPath("O.Next.2"));
    }

    [Fact]
    public void Load_ReadsAttachmentMetadataWithoutPayloads() {
        byte[] payload = Encoding.UTF8.GetBytes("<invoice />");
        byte[] pdf = PdfDocument.Create()
            .AttachFile("invoice.xml", payload, "application/xml", PdfAssociatedFileRelationship.Data, "Invoice XML")
            .Paragraph(p => p.Text("Attachment metadata proof."))
            .ToBytes();

        PdfLogicalDocument logical = PdfLogicalDocument.Load(pdf);

        Assert.True(logical.HasAttachments);
        Assert.Equal(1, logical.AttachmentCount);
        Assert.Equal(new[] { "invoice.xml" }, logical.AttachmentNames);
        Assert.Equal(new[] { "invoice.xml" }, logical.AttachmentFileNames);
        Assert.Equal(new[] { "Names/EmbeddedFiles" }, logical.AttachmentSources);

        PdfAttachmentInfo attachment = Assert.Single(logical.Attachments);
        Assert.Equal("invoice.xml", attachment.Name);
        Assert.Equal("invoice.xml", attachment.FileName);
        Assert.Equal("Invoice XML", attachment.Description);
        Assert.Equal("application/xml", attachment.MimeType);
        Assert.Equal(PdfAssociatedFileRelationship.Data, attachment.Relationship);
        Assert.Equal(payload.Length, attachment.SizeBytes);
        Assert.Same(attachment, Assert.Single(logical.GetAttachmentsByFileName("invoice.xml")));
        Assert.Same(attachment, Assert.Single(logical.GetAttachmentsBySource("Names/EmbeddedFiles")));
        Assert.Same(attachment, Assert.Single(logical.GetAttachmentsByRelationship(PdfAssociatedFileRelationship.Data)));
        Assert.Empty(logical.GetAttachmentsBySource("AF"));
    }

    [Fact]
    public void Load_ReadsPageGeometryAndPresentationMetadata() {
        PdfLogicalDocument logical = PdfLogicalDocument.Load(PdfPageGeometrySupport.BuildPageGeometryPdf());

        PdfLogicalPage page = Assert.Single(logical.Pages);
        Assert.Equal(380, page.Width);
        Assert.Equal(260, page.Height);
        Assert.Equal(400, page.MediaBox!.Width);
        Assert.Equal(10, page.CropBox!.Left);
        Assert.Equal(5, page.BleedBox!.Left);
        Assert.Equal(20, page.TrimBox!.Left);
        Assert.Equal(25, page.ArtBox!.Left);
        Assert.Equal(2, page.UserUnit);
        Assert.Equal("S", page.TabOrder);
        Assert.Equal(5, page.DurationSeconds);
        Assert.True(page.Geometry.HasTransition);
        Assert.Equal("Fly", page.Transition!.Style);
        Assert.Equal(1.5, page.Transition.DurationSeconds);
        Assert.Equal(90, page.Transition.Direction);
        Assert.True(page.HasPageMetadata);
        Assert.True(page.HasPieceInfo);
    }

    [Fact]
    public void Load_ReadsOptionalContentLayerMetadata() {
        PdfLogicalDocument logical = PdfLogicalDocument.Load(PdfOptionalContentSupport.BuildOptionalContentMetadataPdf());

        Assert.True(logical.HasReadableOptionalContent);
        Assert.True(logical.HasOptionalContentGroups);
        Assert.Equal(2, logical.OptionalContentGroupCount);
        Assert.Equal(new[] { "Print layer", "Hidden layer" }, logical.OptionalContentGroupNames);
        Assert.Equal("Default layers", logical.OptionalContent!.DefaultConfigurationName);
        Assert.Equal("ON", logical.OptionalContent.BaseState);

        PdfOptionalContentGroup printLayer = Assert.Single(logical.GetOptionalContentGroupsByName("Print layer"));
        Assert.True(printLayer.IsInitiallyVisible);
        Assert.False(printLayer.IsLocked);
        Assert.Equal(new[] { "View", "Design" }, printLayer.Intents);
        Assert.Equal("OFF", printLayer.ExportState);

        PdfOptionalContentGroup hiddenLayer = Assert.Single(logical.GetOptionalContentGroupsByName("Hidden layer"));
        Assert.False(hiddenLayer.IsInitiallyVisible);
        Assert.True(hiddenLayer.IsLocked);
        Assert.Equal("ON", hiddenLayer.ExportState);
        Assert.Empty(logical.GetOptionalContentGroupsByName("Missing"));
    }
}
