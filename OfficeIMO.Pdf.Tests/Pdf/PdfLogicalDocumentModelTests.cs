using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentReadResultTests {
    [Fact]
    public void PdfLogicalElementKind_PreservesPublishedNumericValues() {
        Assert.Equal(0, (int)PdfLogicalElementKind.TextBlock);
        Assert.Equal(1, (int)PdfLogicalElementKind.Heading);
        Assert.Equal(2, (int)PdfLogicalElementKind.ListItem);
        Assert.Equal(3, (int)PdfLogicalElementKind.LeaderRow);
        Assert.Equal(4, (int)PdfLogicalElementKind.Table);
        Assert.Equal(5, (int)PdfLogicalElementKind.Image);
        Assert.Equal(6, (int)PdfLogicalElementKind.LinkAnnotation);
        Assert.Equal(7, (int)PdfLogicalElementKind.FormWidget);
        Assert.Equal(8, (int)PdfLogicalElementKind.Header);
        Assert.Equal(9, (int)PdfLogicalElementKind.Footer);
        Assert.Equal(10, (int)PdfLogicalElementKind.Caption);
        Assert.Equal(11, (int)PdfLogicalElementKind.Footnote);
    }

    [Fact]
    public void DocumentHeadingTiers_ExcludeAuthoritativeSemanticLevelsFromFontRanking() {
        var explicitBlock = new PdfLogicalTextBlock(
            1, PdfLogicalElementKind.Heading, "Tagged heading", 50D, 250D, 700D, 30D, Array.Empty<PdfTextSpan>());
        var heuristicBlock = new PdfLogicalTextBlock(
            1, PdfLogicalElementKind.Heading, "Heuristic heading", 50D, 250D, 650D, 20D, Array.Empty<PdfTextSpan>());
        var explicitHeading = new PdfLogicalHeading(
            1,
            1,
            explicitBlock.Text,
            explicitBlock.FontSize,
            explicitBlock,
            evidence: new[] {
                new PdfInferenceEvidence("semantic.tagged-pdf-role", "The tagged PDF supplies the heading level.", 1D)
            });
        var heuristicHeading = new PdfLogicalHeading(
            1,
            3,
            heuristicBlock.Text,
            heuristicBlock.FontSize,
            heuristicBlock);

        PdfDocumentReadResult.ApplyDocumentHeadingFontTiers(
            new[] { explicitHeading, heuristicHeading },
            System.Threading.CancellationToken.None);

        Assert.Equal(1, explicitHeading.Level);
        Assert.Equal(1, heuristicHeading.Level);
    }

    [Fact]
    public void DocumentHeadingTiers_RankDistinctSizesAndClusterNearbySizes() {
        double[] fontSizes = { 30D, 29.6D, 28D, 27D, 26D, 25D, 24D, 23D };
        PdfLogicalHeading[] headings = fontSizes.Select((fontSize, index) => {
            var block = new PdfLogicalTextBlock(
                1,
                PdfLogicalElementKind.Heading,
                "Heading " + index,
                50D,
                250D,
                700D - index * 30D,
                fontSize,
                Array.Empty<PdfTextSpan>());
            return new PdfLogicalHeading(1, 1, block.Text, fontSize, block);
        }).ToArray();

        PdfDocumentReadResult.ApplyDocumentHeadingFontTiers(
            headings,
            System.Threading.CancellationToken.None);

        Assert.Equal(new[] { 1, 1, 2, 3, 4, 5, 6, 6 }, headings.Select(static heading => heading.Level));
    }

    [Fact]
    public void Read_BuildsHeadingHierarchyAndDirectSectionOwnership() {
        byte[] pdf = PdfDocument.Create()
            .H1("Operations")
            .Paragraph(paragraph => paragraph.Text("Operations overview."))
            .H2("North region")
            .Paragraph(paragraph => paragraph.Text("North region details."))
            .H1("Finance")
            .Paragraph(paragraph => paragraph.Text("Finance overview."))
            .ToBytes();

        PdfDocumentReadResult result = PdfDocument.Load(pdf).Read();

        Assert.Equal(3, result.AllSections.Count);
        Assert.Equal(2, result.Sections.Count);
        PdfLogicalSection operations = result.Sections[0];
        PdfLogicalSection north = Assert.Single(operations.Children);
        PdfLogicalSection finance = result.Sections[1];
        Assert.Equal("Operations", operations.Title);
        Assert.Equal("North region", north.Title);
        Assert.Same(operations, north.Parent);
        Assert.Equal("Finance", finance.Title);
        Assert.Null(finance.Parent);
        Assert.Contains(operations.Paragraphs, paragraph => paragraph.Text.Contains("Operations overview", StringComparison.Ordinal));
        Assert.Contains(north.Paragraphs, paragraph => paragraph.Text.Contains("North region details", StringComparison.Ordinal));
        Assert.Contains(finance.Paragraphs, paragraph => paragraph.Text.Contains("Finance overview", StringComparison.Ordinal));
        Assert.Same(north, result.GetOwningSection(Assert.Single(north.Paragraphs)));
    }

    [Fact]
    public async System.Threading.Tasks.Task Sections_PublishCompleteOwnershipIndexesDuringConcurrentFirstAccess() {
        byte[] pdf = PdfDocument.Create()
            .H1("Operations")
            .Paragraph(paragraph => paragraph.Text("Operations overview."))
            .H2("North region")
            .Paragraph(paragraph => paragraph.Text("North region details."))
            .ToBytes();
        PdfDocumentReadResult result = PdfDocument.Load(pdf).Read();
        PdfLogicalParagraph paragraph = result.Paragraphs.Last();
        using var start = new System.Threading.ManualResetEventSlim(false);
        System.Threading.Tasks.Task<(IReadOnlyList<PdfLogicalSection> Roots, IReadOnlyList<PdfLogicalSection> All, PdfLogicalSection? Owner)>[] readers =
            Enumerable.Range(0, 32)
                .Select(_ => System.Threading.Tasks.Task.Run(() => {
                    start.Wait();
                    return (result.Sections, result.AllSections, result.GetOwningSection(paragraph));
                }))
                .ToArray();

        start.Set();
        (IReadOnlyList<PdfLogicalSection> Roots, IReadOnlyList<PdfLogicalSection> All, PdfLogicalSection? Owner)[] snapshots =
            await System.Threading.Tasks.Task.WhenAll(readers);

        Assert.All(snapshots, snapshot => {
            Assert.Single(snapshot.Roots);
            Assert.Equal(2, snapshot.All.Count);
            Assert.Same(snapshot.All[1], snapshot.Owner);
        });
    }

    [Fact]
    public void Read_AssignsLinksAndFormWidgetsToHeadingOwnedSections() {
        byte[] source = PdfDocument.Create()
            .H1("Interactive section")
            .Paragraph(paragraph => paragraph.Link("Section link", "https://example.com/section"))
            .ToBytes();
        PdfDocument document = PdfDocument.Load(source).Forms.Edit(edit => edit.Create(new PdfFormFieldCreateOptions {
            Name = "section-note",
            Kind = PdfFormFieldCreationKind.Text,
            PageNumber = 1,
            X = 72,
            Y = 420,
            Width = 180,
            Height = 24
        })).ToDocument();

        PdfDocumentReadResult result = document.Read();
        PdfLogicalSection section = Assert.Single(result.Sections);
        PdfLogicalLinkAnnotation link = Assert.Single(result.Links);
        PdfLogicalFormWidget widget = Assert.Single(result.FormWidgets);

        Assert.Contains(link, section.Links);
        Assert.Contains(widget, section.FormWidgets);
        Assert.Same(section, result.GetOwningSection(link));
        Assert.Same(section, result.GetOwningSection(widget));
    }

    [Fact]
    public void Read_CreatesOccurrenceLocalWidgetsForRepeatedPageSelections() {
        byte[] source = PdfDocument.Create()
            .H1("Repeated interactive section")
            .Paragraph(paragraph => paragraph.Text("Repeated page body"))
            .ToBytes();
        PdfDocument document = PdfDocument.Load(source).Forms.Edit(edit => edit.Create(new PdfFormFieldCreateOptions {
            Name = "repeated-note",
            Kind = PdfFormFieldCreationKind.Text,
            PageNumber = 1,
            X = 72,
            Y = 420,
            Width = 180,
            Height = 24
        })).ToDocument();

        PdfDocumentReadResult result = document.Read(new PdfReadOptions {
            PageSelection = PdfPageSelection.From(1, 1)
        });
        PdfLogicalFormWidget first = Assert.Single(result.Pages[0].FormWidgets);
        PdfLogicalFormWidget second = Assert.Single(result.Pages[1].FormWidgets);
        PdfLogicalSection firstOwner = Assert.IsType<PdfLogicalSection>(result.GetOwningSection(first));
        PdfLogicalSection secondOwner = Assert.IsType<PdfLogicalSection>(result.GetOwningSection(second));

        Assert.NotSame(first, second);
        Assert.NotSame(firstOwner, secondOwner);
        Assert.Contains(first, firstOwner.FormWidgets);
        Assert.DoesNotContain(second, firstOwner.FormWidgets);
        Assert.Contains(second, secondOwner.FormWidgets);
        Assert.DoesNotContain(first, secondOwner.FormWidgets);
    }

    [Fact]
    public void Read_TracksRepeatedImagePlacementsWithSectionLocalOwnership() {
        byte[] image = PdfPngTestImages.CreateRgbPng(20, 80, 160);
        byte[] source = PdfDocument.Create()
            .H1("First image section")
            .Paragraph(paragraph => paragraph.Text("First section body"))
            .Image(image, 24, 24)
            .H1("Second image section")
            .Paragraph(paragraph => paragraph.Text("Second section body"))
            .Image(image, 24, 24)
            .ToBytes();

        PdfDocumentReadResult result = PdfDocument.Load(source).Read();
        PdfLogicalImage resource = Assert.Single(Assert.Single(result.Pages).Images);
        Assert.Equal(2, resource.PlacementCount);
        Assert.Equal(2, result.Sections.Count);

        PdfLogicalImage firstPlacement = Assert.Single(result.Sections[0].Images);
        PdfLogicalImage secondPlacement = Assert.Single(result.Sections[1].Images);
        Assert.NotSame(firstPlacement, secondPlacement);
        Assert.Same(resource.SourceImage, firstPlacement.SourceImage);
        Assert.Same(resource.SourceImage, secondPlacement.SourceImage);
        Assert.Single(firstPlacement.Placements);
        Assert.Single(secondPlacement.Placements);
        Assert.NotSame(firstPlacement.PrimaryPlacement, secondPlacement.PrimaryPlacement);
        Assert.Same(result.Sections[0], result.GetOwningSection(firstPlacement));
        Assert.Same(result.Sections[1], result.GetOwningSection(secondPlacement));
        Assert.Null(result.GetOwningSection(resource));
        Assert.Equal(result.Sections, result.GetOwningSections(resource));
    }

    [Fact]
    public void Read_DoesNotAssignAggregateImageToASectionWhenAnotherPlacementIsUnsectioned() {
        byte[] image = PdfPngTestImages.CreateRgbPng(160, 80, 20);
        byte[] source = PdfDocument.Create()
            .Image(image, 24, 24)
            .H1("Owned image section")
            .Paragraph(paragraph => paragraph.Text("Owned section body"))
            .Image(image, 24, 24)
            .ToBytes();

        PdfDocumentReadResult result = PdfDocument.Load(source).Read();
        PdfLogicalImage resource = Assert.Single(Assert.Single(result.Pages).Images);
        PdfLogicalSection section = Assert.Single(result.Sections);
        PdfLogicalImage placement = Assert.Single(section.Images);

        Assert.Equal(2, resource.PlacementCount);
        Assert.Same(section, result.GetOwningSection(placement));
        Assert.Null(result.GetOwningSection(resource));
        Assert.Equal(new[] { section }, result.GetOwningSections(resource));
    }

    [Fact]
    public void Read_AssignsAggregateImageWhenEveryPlacementSharesOneSection() {
        byte[] image = PdfPngTestImages.CreateRgbPng(80, 160, 20);
        byte[] source = PdfDocument.Create()
            .H1("Shared image section")
            .Paragraph(paragraph => paragraph.Text("Shared section body"))
            .Image(image, 24, 24)
            .Image(image, 24, 24)
            .ToBytes();

        PdfDocumentReadResult result = PdfDocument.Load(source).Read();
        PdfLogicalImage resource = Assert.Single(Assert.Single(result.Pages).Images);
        PdfLogicalSection section = Assert.Single(result.Sections);

        Assert.Equal(2, resource.PlacementCount);
        Assert.Equal(2, section.Images.Count);
        Assert.Same(section, result.GetOwningSection(resource));
        Assert.Equal(new[] { section }, result.GetOwningSections(resource));
    }

    [Fact]
    public void LogicalTextBlocks_PreservePositionedRunStyleSpans() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph
                .Color(PdfColor.FromRgb(255, 0, 0))
                .Font(PdfStandardFont.Helvetica)
                .Text("Red")
                .Color(PdfColor.FromRgb(0, 0, 255))
                .Font(PdfStandardFont.Courier)
                .Text("Blue"))
            .ToBytes();

        PdfLogicalTextBlock block = Assert.Single(PdfDocumentReadResult.Load(pdf).TextBlocks);

        Assert.Equal(block.SpanCount, block.Spans.Count);
        Assert.True(block.Spans.Count >= 2);
        Assert.Contains(block.Spans, span => span.Text.Contains("Red", StringComparison.Ordinal));
        Assert.Contains(block.Spans, span => span.Text.Contains("Blue", StringComparison.Ordinal));
        Assert.Contains(block.Spans, span => span.BaseFont?.Contains("Helvetica", StringComparison.OrdinalIgnoreCase) == true);
        Assert.Contains(block.Spans, span => span.BaseFont?.Contains("Courier", StringComparison.OrdinalIgnoreCase) == true);
        Assert.True(block.Spans.Where(span => span.Color.HasValue).Select(span => span.Color).Distinct().Count() >= 2);
        Assert.Equal(block.Text, string.Concat(block.Runs.Select(run => run.Text)));
        Assert.Contains(block.Runs, run =>
            run.Text.Contains("Red", StringComparison.Ordinal) &&
            run.BaseFont?.Contains("Helvetica", StringComparison.OrdinalIgnoreCase) == true);
        Assert.Contains(block.Runs, run =>
            run.Text.Contains("Blue", StringComparison.Ordinal) &&
            run.BaseFont?.Contains("Courier", StringComparison.OrdinalIgnoreCase) == true);
        Assert.True(block.Runs.Where(run => run.Color.HasValue).Select(run => run.Color).Distinct().Count() >= 2);
    }

    [Fact]
    public void LogicalTextBlocks_AlignRunsAfterWhitespaceNormalization() {
        var first = new PdfTextSpan(
            "  Red  ",
            "F1",
            12,
            0,
            100,
            30,
            OfficeIMO.Drawing.OfficeColor.Red,
            baseFont: "Helvetica");
        var second = new PdfTextSpan(
            "Blue",
            "F2",
            12,
            36,
            100,
            24,
            OfficeIMO.Drawing.OfficeColor.Blue,
            baseFont: "Courier");

        var block = new PdfLogicalTextBlock(
            1,
            PdfLogicalElementKind.TextBlock,
            "Red Blue",
            0,
            60,
            100,
            12,
            new[] { first, second });

        Assert.Equal("Red Blue", string.Concat(block.Runs.Select(run => run.Text)));
        Assert.Collection(
            block.Runs,
            run => {
                Assert.Equal("Red ", run.Text);
                Assert.Equal("Helvetica", run.BaseFont);
                Assert.Equal(OfficeIMO.Drawing.OfficeColor.Red, run.Color);
            },
            run => {
                Assert.Equal("Blue", run.Text);
                Assert.Equal("Courier", run.BaseFont);
                Assert.Equal(OfficeIMO.Drawing.OfficeColor.Blue, run.Color);
            });
    }

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

        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(pdf, new PdfTextLayoutOptions {
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
        Assert.InRange(heading.Confidence, 0D, 1D);
        Assert.NotEmpty(heading.Evidence);
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
        Assert.InRange(listItem.Confidence, 0D, 1D);
        Assert.NotEmpty(listItem.Evidence);
        Assert.Same(listItem, Assert.Single(logical.ListItems));
        Assert.Contains(page.Paragraphs, paragraph => Normalize(paragraph.Text).Contains("Logicalreadbackmarker", StringComparison.Ordinal));
        Assert.Contains(logical.Paragraphs, paragraph => Normalize(paragraph.Text).Contains("Logicalreadbackmarker", StringComparison.Ordinal));
        Assert.DoesNotContain(page.Paragraphs, paragraph => Normalize(paragraph.Text).Contains("A-100", StringComparison.Ordinal));

        PdfLogicalTable table = Assert.Single(page.Tables, item => item.Rows.Count >= 3 && item.Columns.Count >= 3);
        Assert.InRange(table.Confidence, 0D, 1D);
        Assert.NotEmpty(table.Evidence);
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

        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(pdf);

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

        PdfDocumentReadResult pageRange = PdfDocumentReadResult.LoadPageRanges(pdf, new PdfPageRange(1, 1));
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

        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(pdf);

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

        PdfDocumentReadResult pageRange = PdfDocumentReadResult.LoadPageRanges(pdf, new PdfPageRange(1, 1));
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

        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(pdf);

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

        PdfDocumentReadResult pageRange = PdfDocumentReadResult.LoadPageRanges(pdf, new PdfPageRange(1, 1));
        Assert.False(pageRange.HasReadableTaggedContent);
        Assert.Null(pageRange.TaggedContent);
    }

    [Fact]
    public void Load_ReadsCatalogActionsWithoutScriptPayload() {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(BuildCatalogJavaScriptActionPdf());

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
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(BuildCatalogJavaScriptActionWithSharedNextActionPdf());

        Assert.True(logical.HasCatalogActions);
        Assert.Equal(3, logical.CatalogActionCount);
        Assert.Equal(new[] { "OpenAction", "OpenAction.Next.0", "OpenAction.Next.1" }, logical.CatalogActionNames);
        Assert.Equal(2, logical.GetCatalogActionsByActionType("Launch").Count);
    }

    [Fact]
    public void Load_ReadsPageActionsWithoutScriptPayload() {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(BuildPageAdditionalActionsPdf());

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
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(BuildPageAdditionalActionsWithSharedIndirectActionPdf());

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
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(BuildPageChainedActionsPdf());

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

        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(pdf);

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
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(PdfPageGeometrySupport.BuildPageGeometryPdf());

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
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(PdfOptionalContentSupport.BuildOptionalContentMetadataPdf());

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

    [Fact]
    public void Load_DoesNotRemoveDecimalParagraphsAsListItems() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("1037.25 total"))
            .Paragraph(paragraph => paragraph.Text("-42 total"))
            .Paragraph(paragraph => paragraph.Text("-$42 total"))
            .Paragraph(paragraph => paragraph.Text("-.5 variance"))
            .Paragraph(paragraph => paragraph.Text("1. Actual numbered item"))
            .Paragraph(paragraph => paragraph.Text("2)Compact numbered item"))
            .Paragraph(paragraph => paragraph.Text("(a)Compact parenthesized item"))
            .Paragraph(paragraph => paragraph.Text("-Compact ASCII bullet"))
            .Paragraph(paragraph => paragraph.Text("*Compact starred bullet"))
            .ToBytes();

        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(pdf);

        Assert.DoesNotContain(logical.ListItems, item => item.Text.Contains("1037.25", StringComparison.Ordinal));
        Assert.DoesNotContain(logical.ListItems, item => item.Text.Contains("-42", StringComparison.Ordinal));
        Assert.DoesNotContain(logical.ListItems, item => item.Text.Contains("-$42", StringComparison.Ordinal));
        Assert.DoesNotContain(logical.ListItems, item => item.Text.Contains("-.5", StringComparison.Ordinal));
        Assert.Contains(logical.Paragraphs, paragraph => paragraph.Text.Contains("1037.25 total", StringComparison.Ordinal));
        Assert.Contains(logical.Paragraphs, paragraph => paragraph.Text.Contains("-42 total", StringComparison.Ordinal));
        Assert.Contains(logical.Paragraphs, paragraph => paragraph.Text.Contains("-$42 total", StringComparison.Ordinal));
        Assert.Contains(logical.Paragraphs, paragraph => paragraph.Text.Contains("-.5 variance", StringComparison.Ordinal));
        Assert.Contains(logical.ListItems, item => item.Text.Contains("Actual numbered item", StringComparison.Ordinal));
        Assert.Contains(logical.ListItems, item => item.Text.Contains("Compact numbered item", StringComparison.Ordinal));
        Assert.Contains(logical.ListItems, item => item.Text.Contains("Compact parenthesized item", StringComparison.Ordinal));
        Assert.Contains(logical.ListItems, item => item.Text.Contains("Compact ASCII bullet", StringComparison.Ordinal));
        Assert.Contains(logical.ListItems, item => item.Text.Contains("Compact starred bullet", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("٣) عنصر", "٣", "عنصر")]
    [InlineData("(甲) 项目", "(甲)", "项目")]
    [InlineData("（甲）項目", "（甲）", "項目")]
    [InlineData("３、項目", "３", "項目")]
    [InlineData("一、項目", "一", "項目")]
    [InlineData("◦ punkt", "◦", "punkt")]
    [InlineData("‣ элемент", "‣", "элемент")]
    public void ListSyntax_UsesUnicodeMarkersWithoutLanguageVocabulary(string source, string expectedMarker, string expectedBody) {
        bool parsed = ContentStructureExtractor.TryParseListItemText(
            source,
            out string marker,
            out string body,
            out int level);

        Assert.True(parsed);
        Assert.Equal(expectedMarker, marker);
        Assert.Equal(expectedBody, body);
        Assert.Equal(1, level);
    }

    [Theory]
    [InlineData("(Appendix) text")]
    [InlineData("(Important) text")]
    public void ListSyntax_DoesNotTreatParenthesizedWordsAsMarkers(string source) {
        Assert.False(ContentStructureExtractor.IsListItemText(source));
    }

    [Fact]
    public void TextNormalization_DoesNotJoinLexicalFragmentsWithoutGeometry() {
        Assert.Equal("inform ation", ContentStructureExtractor.NormalizeShattered("inform ation"));
        Assert.Equal("för säkring", ContentStructureExtractor.NormalizeShattered("för säkring"));
        Assert.Equal("文 書", ContentStructureExtractor.NormalizeShattered("文 書"));
    }

    [Fact]
    public void TextEmission_PreservesVisibleHyphensAndOnlyRejoinsSoftHyphensWhenEnabled() {
        var visibleLines = new List<TextLayoutEngine.TextLine> {
            new(700D, 50D, 120D, "multi-", new List<PdfTextSpan>()),
            new(680D, 50D, 120D, "lingual", new List<PdfTextSpan>())
        };
        var softLines = new List<TextLayoutEngine.TextLine> {
            new(700D, 50D, 120D, "multi\u00AD", new List<PdfTextSpan>()),
            new(680D, 50D, 120D, "lingual", new List<PdfTextSpan>())
        };
        var columns = new TextLayoutEngine.ColumnLayout((0D, 500D), (0D, 0D), false);

        Assert.Equal("multi-\nlingual", TextLayoutEngine.EmitText(visibleLines, columns));
        Assert.Equal(
            "multi-\nlingual",
            TextLayoutEngine.EmitText(visibleLines, columns, new PdfTextLayoutOptions { JoinSoftHyphensAcrossLines = true }));
        Assert.Equal("multi\u00AD\nlingual", TextLayoutEngine.EmitText(softLines, columns));
        Assert.Equal(
            "multilingual",
            TextLayoutEngine.EmitText(softLines, columns, new PdfTextLayoutOptions { JoinSoftHyphensAcrossLines = true }));
    }
}
