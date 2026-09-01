using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Tests.Pdf;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public class PowerPointPdfTableImportTests {
    [Fact]
    public void PdfDocument_ToPowerPointPresentation_VisualProfileCreatesOnePageImagePerSlide() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 360,
                PageHeight = 240,
                MarginLeft = 24,
                MarginRight = 24,
                MarginTop = 24,
                MarginBottom = 24
            })
            .H1("Visual page one")
            .PageBreak()
            .Paragraph(p => p.Text("Visual page two"))
            .ToBytes();

        PdfCore.PdfDocument opened = PdfCore.PdfDocument.Load(pdf);
        PdfPowerPointConversionResult result = opened.ToPowerPointPresentationResult(
            PdfPowerPointImportOptions.CreateVisualPages());

        Assert.Equal(PdfPowerPointImportMode.VisualPages, result.Report.Mode);
        Assert.Equal(new[] { 1, 2 }, result.Report.VisualPages.Select(page => page.PageNumber).ToArray());
        Assert.All(result.Report.VisualPages, page => Assert.True(page.Succeeded));
        Assert.Empty(result.Report.TableEntries);
        Assert.Contains(result.Warnings, warning => warning.Code == "PdfVisualPageSlidesNotEditable");

        using var presentation = new MemoryStream();
        using (result.Value) {
            result.Value.Save(presentation);
        }

        using PresentationDocument package = PresentationDocument.Open(new MemoryStream(presentation.ToArray()), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
        SlidePart[] slideParts = package.PresentationPart!.SlideParts.ToArray();
        Assert.Equal(2, slideParts.Length);
        Assert.All(
            slideParts,
            slidePart => Assert.Single(slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Picture>()));
    }

    [Fact]
    public void PdfDocument_ToPowerPointPresentation_DefaultCreatesNativeEditableObjects() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .H1("Quarterly report")
            .Paragraph(p => p.Text("Editable content projection"))
            .Image(PdfPngTestImages.CreateRgbPng(2, 2), 18, 18, alternativeText: "Status marker")
            .Table(new[] {
                new[] { "Metric", "Value" },
                new[] { "Ready", "Yes" },
                new[] { "Validated", "Yes" }
            }, style: new PdfCore.PdfTableStyle {
                HeaderRowCount = 1,
                ColumnWidthPoints = new List<double?> { 160, 90 }
            })
            .ToBytes();

        Assert.Equal(PdfPowerPointImportMode.Auto, new PdfPowerPointImportOptions().Mode);

        PdfPowerPointConversionResult result = PdfCore.PdfDocument.Load(pdf)
            .ToPowerPointPresentationResult();

        Assert.Equal(PdfPowerPointImportMode.EditableContent, result.Report.Mode);
        PdfPowerPointEditablePageEntry page = Assert.Single(result.Report.EditablePages);
        Assert.True(page.TextBoxCount >= 2);
        Assert.Equal(1, page.TableCount);
        Assert.Equal(1, page.ImageCount);
        Assert.Single(result.Report.TableEntries);
        Assert.Contains(result.Warnings, warning => warning.Code == "PdfEditableContentReconstructed");

        using var presentation = new MemoryStream();
        using (result.Value) result.Value.Save(presentation);
        using PresentationDocument package = PresentationDocument.Open(new MemoryStream(presentation.ToArray()), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
        SlidePart slide = Assert.Single(package.PresentationPart!.SlideParts);
        Assert.Single(slide.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Picture>());
        Assert.Single(slide.Slide.Descendants<A.Table>());
        Assert.Contains(slide.Slide.Descendants<A.Text>(), text => text.Text == "Quarterly report");
        Assert.Contains(slide.Slide.Descendants<A.Text>(), text => text.Text == "Ready");
    }

    [Fact]
    public void PdfDocument_ToPowerPointPresentation_EditableContentCreatesNativeSafeVectorShapes() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 420,
                PageHeight = 300,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                PageBackgroundShapes = new[] {
                    PdfCore.PdfPageBackgroundShape.Rectangle(
                        40,
                        210,
                        120,
                        36,
                        fill: PdfCore.PdfColor.FromRgb(219, 234, 254),
                        stroke: PdfCore.PdfColor.FromRgb(37, 99, 235),
                        strokeWidth: 1)
                }
            })
            .Paragraph(p => p.Text("Editable rectangle"))
            .ToBytes();

        PdfPowerPointConversionResult result = PdfCore.PdfDocument.Load(pdf)
            .ToPowerPointPresentationResult(PdfPowerPointImportOptions.CreateEditableContent());

        PdfPowerPointEditablePageEntry page = Assert.Single(result.Report.EditablePages);
        Assert.True(page.ShapeCount >= 1);
        using var presentation = new MemoryStream();
        using (result.Value) result.Value.Save(presentation);
        using PresentationDocument package = PresentationDocument.Open(new MemoryStream(presentation.ToArray()), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
        SlidePart slide = Assert.Single(package.PresentationPart!.SlideParts);
        Assert.Contains(
            slide.Slide.Descendants<A.PresetGeometry>(),
            geometry => geometry.Preset?.Value == A.ShapeTypeValues.Rectangle);
    }

    [Theory]
    [InlineData(PdfPowerPointImportMode.VisualPages)]
    [InlineData(PdfPowerPointImportMode.EditableTables)]
    [InlineData(PdfPowerPointImportMode.HybridVisualAndEditableTables)]
    [InlineData(PdfPowerPointImportMode.EditableContent)]
    [InlineData(PdfPowerPointImportMode.Auto)]
    public void PdfDocument_ToPowerPointPresentation_EnforcesPageLimitInEveryMode(
        PdfPowerPointImportMode mode) {
        byte[] pdf = PdfCore.PdfDocument.Create()
            .Paragraph(p => p.Text("Page one"))
            .PageBreak()
            .Paragraph(p => p.Text("Page two"))
            .ToBytes();
        var options = new PdfPowerPointImportOptions {
            Mode = mode,
            MaxPages = 1
        };

        Exception exception = Assert.ThrowsAny<Exception>(() =>
            PdfCore.PdfDocument.Load(pdf).ToPowerPointPresentationResult(options));

        Assert.Contains("limit", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(PdfPowerPointImportMode.EditableTables)]
    [InlineData(PdfPowerPointImportMode.EditableContent)]
    public void PdfDocument_ToPowerPointPresentation_RejectsOversizedSelectionBeforeLogicalExtraction(
        PdfPowerPointImportMode mode) {
        byte[] pdf = PdfCore.PdfDocument.Create()
            .Paragraph(p => p.Text("Only one source page"))
            .ToBytes();
        var options = new PdfPowerPointImportOptions {
            Mode = mode,
            MaxPages = 1,
            ReadOptions = new PdfCore.PdfReadOptions {
                PageSelection = PdfCore.PdfPageSelection.From(2, 2)
            }
        };

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            PdfCore.PdfDocument.Load(pdf).ToPowerPointPresentationResult(options));

        Assert.Contains("page count 2", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("configured limit of 1", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void PdfDocument_ToPowerPointPresentation_AllowsRaisedDestinationAndSemanticPageLimits() {
        const int selectedPageCount = 1_001;
        byte[] pdf = PdfCore.PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Repeated source page"))
            .ToBytes();
        var options = new PdfPowerPointImportOptions {
            Mode = PdfPowerPointImportMode.EditableTables,
            MaxPages = selectedPageCount,
            ReadOptions = new PdfCore.PdfReadOptions {
                PageSelection = PdfCore.PdfPageSelection.From(Enumerable.Repeat(1, selectedPageCount).ToArray()),
                Pipeline = new PdfCore.PdfUnderstandingPipelineOptions { MaxPages = selectedPageCount }
            }
        };

        PdfPowerPointConversionResult result = PdfCore.PdfDocument.Load(pdf)
            .ToPowerPointPresentationResult(options);

        using (result.Value) {
            Assert.Single(result.Value.Slides);
        }
    }

    [Fact]
    public void PdfDocument_ToPowerPointPresentation_EditableObjectLimitReportsTextAndTableLoss() {
        byte[] pdf = PdfCore.PdfDocument.Create()
            .H1("Limit contract")
            .Paragraph(p => p.Text("Second editable text block"))
            .Table(new[] {
                new[] { "Metric", "Value" },
                new[] { "Ready", "Yes" },
                new[] { "Validated", "Yes" }
            }, style: new PdfCore.PdfTableStyle { HeaderRowCount = 1 })
            .ToBytes();
        var options = PdfPowerPointImportOptions.CreateEditableContent();
        options.MaxEditableObjectsPerPage = 1;

        PdfPowerPointConversionResult result = PdfCore.PdfDocument.Load(pdf)
            .ToPowerPointPresentationResult(options);

        using (result.Value) {
            PdfPowerPointEditablePageEntry page = Assert.Single(result.Report.EditablePages);
            Assert.True(page.OmittedTextCount > 0);
            Assert.True(page.OmittedTableCount > 0);
            Assert.True(page.HasOmittedContent);
            Assert.True(result.HasLoss);
            Assert.Contains(result.Warnings, warning => warning.Code == "PdfEditableObjectLimitReached");
            Assert.Throws<InvalidOperationException>(result.Report.RequireNoLoss);
        }
    }

    [Theory]
    [InlineData(360D, 180D, 2000)]
    [InlineData(1440D, 720D, 500)]
    public void PdfDocument_ToPowerPointPresentation_EditableTextScalesWithPageGeometry(
        double pageWidth,
        double pageHeight,
        int expectedFontSizeHundredths) {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = pageWidth,
                PageHeight = pageHeight,
                MarginLeft = 12,
                MarginRight = 12,
                MarginTop = 12,
                MarginBottom = 12,
                DefaultFontSize = 10
            })
            .Paragraph(p => p.Text("Scale probe"))
            .ToBytes();
        PdfPowerPointConversionResult result = PdfCore.PdfDocument.Load(pdf)
            .ToPowerPointPresentationResult(PdfPowerPointImportOptions.CreateEditableContent());

        using var presentation = new MemoryStream();
        using (result.Value) result.Value.Save(presentation);
        using PresentationDocument package = PresentationDocument.Open(new MemoryStream(presentation.ToArray()), false);
        A.Run run = Assert.Single(
            package.PresentationPart!.SlideParts.Single().Slide.Descendants<A.Run>(),
            candidate => candidate.Text?.Text == "Scale probe");
        Assert.Equal(expectedFontSizeHundredths, run.RunProperties?.FontSize?.Value);
    }

    [Theory]
    [InlineData(90, 270)]
    [InlineData(180, 180)]
    [InlineData(270, 90)]
    public void PdfDocument_ToPowerPointPresentation_EditableTextUsesVisualPageRotation(
        int pageRotation,
        int expectedPowerPointRotation) {
        byte[] source = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 360,
                PageHeight = 240,
                MarginLeft = 24,
                MarginRight = 24,
                MarginTop = 24,
                MarginBottom = 24
            })
            .Paragraph(p => p.Text("Rotation probe"))
            .ToBytes();
        byte[] rotated = PdfCore.PdfDocument.Load(source).Pages.Rotate(pageRotation, "1").ToBytes();
        PdfPowerPointConversionResult result = PdfCore.PdfDocument.Load(rotated)
            .ToPowerPointPresentationResult(PdfPowerPointImportOptions.CreateEditableContent());

        using var presentation = new MemoryStream();
        using (result.Value) result.Value.Save(presentation);
        using PresentationDocument package = PresentationDocument.Open(new MemoryStream(presentation.ToArray()), false);
        P.Shape shape = Assert.Single(
            package.PresentationPart!.SlideParts.Single().Slide.Descendants<P.Shape>(),
            candidate => candidate.TextBody?.InnerText == "Rotation probe");
        Assert.Equal(expectedPowerPointRotation * 60000, shape.ShapeProperties?.Transform2D?.Rotation?.Value);
    }

    [Fact]
    public void PdfDocument_ToPowerPointPresentation_EditableContinuationTablesStayInsideSlide() {
        byte[] pdf = PdfCore.PdfDocument.Create()
            .Table(new[] {
                new[] { "Metric", "Value" },
                new[] { "One", "1" },
                new[] { "Two", "2" },
                new[] { "Three", "3" }
            }, style: new PdfCore.PdfTableStyle { HeaderRowCount = 1 })
            .ToBytes();
        var options = PdfPowerPointImportOptions.CreateEditableContent();
        options.MaxRowsPerSlide = 1;
        options.IncludeSourceTitles = false;

        PdfPowerPointConversionResult result = PdfCore.PdfDocument.Load(pdf)
            .ToPowerPointPresentationResult(options);

        using (result.Value) {
            Assert.True(result.Value.Slides.Count > 1);
            double slideWidth = result.Value.SlideSize.WidthPoints;
            double slideHeight = result.Value.SlideSize.HeightPoints;
            foreach (OfficeIMO.PowerPoint.PowerPointSlide slide in result.Value.Slides.Skip(1)) {
                OfficeIMO.PowerPoint.PowerPointTable table = Assert.Single(slide.Tables);
                Assert.InRange(table.LeftPoints, 0D, slideWidth);
                Assert.InRange(table.TopPoints, 0D, slideHeight);
                Assert.True(table.LeftPoints + table.WidthPoints <= slideWidth + 0.01D);
                Assert.True(table.TopPoints + table.HeightPoints <= slideHeight + 0.01D);
            }
        }
    }

    [Fact]
    public void PdfDocument_ToPowerPointPresentation_HybridKeepsVisualPageAndEditableTableOverlay() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Paragraph(p => p.Text("Quarterly status"))
            .Table(new[] {
                new[] { "Code", "Qty" },
                new[] { "A-100", "2" },
                new[] { "B-200", "14" }
            }, style: new PdfCore.PdfTableStyle {
                HeaderRowCount = 1,
                ColumnWidthPoints = new List<double?> { 180, 80 },
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .ToBytes();

        PdfPowerPointConversionResult result = PdfCore.PdfDocument.Load(pdf)
            .ToPowerPointPresentationResult(PdfPowerPointImportOptions.CreateHybrid());

        Assert.Equal(PdfPowerPointImportMode.HybridVisualAndEditableTables, result.Report.Mode);
        Assert.Single(result.Report.VisualPages);
        Assert.Single(result.Report.TableEntries);
        Assert.True(result.Report.VisualPages[0].Succeeded);
        Assert.True(result.Report.HasNonEditablePageContent);
        Assert.False(result.Report.HasOmittedPageContent);
        PdfCore.PdfConversionWarning textWarning = Assert.Single(result.Warnings, warning => warning.Code == "PdfTextNotEditable");
        Assert.Equal(PdfCore.PdfConversionWarningSeverity.Information, textWarning.Severity);
        Assert.Equal("VisualOnly", textWarning.Details["Disposition"]);
        Assert.Contains(result.Warnings, warning => warning.Code == "PdfVectorsNotEditable");

        using var presentation = new MemoryStream();
        using (result.Value) result.Value.Save(presentation);
        using PresentationDocument package = PresentationDocument.Open(new MemoryStream(presentation.ToArray()), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
        SlidePart slide = Assert.Single(package.PresentationPart!.SlideParts);
        Assert.Single(slide.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Picture>());
        Assert.Single(slide.Slide.Descendants<A.Table>());
    }

    [Fact]
    public void PdfDocument_ToPowerPointPresentation_HybridMapsRotatedTableBoundsToVisualCoordinates() {
        byte[] source = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 420,
                PageHeight = 300,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Table(new[] {
                new[] { "Code", "Qty" },
                new[] { "A-100", "2" },
                new[] { "B-200", "14" }
            }, style: new PdfCore.PdfTableStyle {
                HeaderRowCount = 1,
                ColumnWidthPoints = new List<double?> { 140, 80 }
            })
            .ToBytes();
        byte[] rotated = PdfCore.PdfPageEditor.RotatePages(source, 90, 1);

        PdfPowerPointConversionResult result = PdfCore.PdfDocument.Load(rotated)
            .ToPowerPointPresentationResult(PdfPowerPointImportOptions.CreateHybrid());
        using var presentation = new MemoryStream();
        using (result.Value) result.Value.Save(presentation);
        using PresentationDocument package = PresentationDocument.Open(new MemoryStream(presentation.ToArray()), false);
        SlidePart slide = Assert.Single(package.PresentationPart!.SlideParts);
        DocumentFormat.OpenXml.Presentation.Picture picture = Assert.Single(
            slide.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Picture>());
        DocumentFormat.OpenXml.Presentation.GraphicFrame frame = Assert.Single(
            slide.Slide.Descendants<DocumentFormat.OpenXml.Presentation.GraphicFrame>());
        long pictureTop = picture.ShapeProperties!.Transform2D!.Offset!.Y!.Value;
        long pictureHeight = picture.ShapeProperties.Transform2D.Extents!.Cy!.Value;
        long tableTop = frame.Transform!.Offset!.Y!.Value;

        Assert.True(tableTop > pictureTop + pictureHeight / 2L);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
    }

    [Fact]
    public void PdfDocument_ToPowerPointPresentation_HybridOmitsSyntheticKeyValueHeader() {
        byte[] pdf = PdfCore.PdfDocument.Create()
            .KeyValueTable(new[] {
                PdfCore.PdfKeyValueRow.Text("InvoiceId", "INV-001"),
                PdfCore.PdfKeyValueRow.Text("Customer", "Evotec"),
                PdfCore.PdfKeyValueRow.Text("Status", "Open")
            }, includeHeader: false)
            .ToBytes();

        PdfPowerPointConversionResult result = PdfCore.PdfDocument.Load(pdf)
            .ToPowerPointPresentationResult(PdfPowerPointImportOptions.CreateHybrid());
        using var presentation = new MemoryStream();
        using (result.Value) result.Value.Save(presentation);
        using PresentationDocument package = PresentationDocument.Open(new MemoryStream(presentation.ToArray()), false);
        A.Table table = GetSingleTable(package);

        Assert.True(ContainsRows(
            table,
            new[] { "InvoiceId", "INV-001" },
            new[] { "Customer", "Evotec" },
            new[] { "Status", "Open" }));
        Assert.DoesNotContain(table.Descendants<A.Text>(), text => text.Text == "Key" || text.Text == "Value");
    }

    [Fact]
    public void PdfDocument_ToPowerPointPresentation_HybridPreservesSelectedPageIndexesInTableReports() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 420,
                PageHeight = 300,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Table(new[] {
                new[] { "Code", "Qty" },
                new[] { "A-100", "2" },
                new[] { "B-200", "14" }
            }, style: new PdfCore.PdfTableStyle { HeaderRowCount = 1, ColumnWidthPoints = new List<double?> { 140, 80 } })
            .PageBreak()
            .Table(new[] {
                new[] { "Name", "Total" },
                new[] { "Alpha", "20" },
                new[] { "Beta", "30" }
            }, style: new PdfCore.PdfTableStyle { HeaderRowCount = 1, ColumnWidthPoints = new List<double?> { 140, 80 } })
            .ToBytes();

        PdfPowerPointConversionResult result = PdfCore.PdfDocument.Load(pdf)
            .ToPowerPointPresentationResult(PdfPowerPointImportOptions.CreateHybrid());

        Assert.Equal(new[] { 0, 1 }, result.Report.TableEntries.Select(entry => entry.PageIndex).ToArray());
        Assert.Equal(new[] { 1, 2 }, result.Report.TableEntries.Select(entry => entry.PageNumber).ToArray());
        result.Value.Dispose();
    }

    [Fact]
    public void PdfDocument_ToPowerPointPresentation_HybridSplitsTablesWithinPerSlideCaps() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 520,
                PageHeight = 420,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 9
            })
            .Table(new[] {
                new[] { "C1", "C2", "C3", "C4" },
                new[] { "R1C1", "R1C2", "R1C3", "R1C4" },
                new[] { "R2C1", "R2C2", "R2C3", "R2C4" },
                new[] { "R3C1", "R3C2", "R3C3", "R3C4" },
                new[] { "R4C1", "R4C2", "R4C3", "R4C4" }
            }, style: new PdfCore.PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 80, 80, 80, 80 },
                HeaderRowCount = 1,
                CellPaddingX = 4,
                CellPaddingY = 3
            })
            .ToBytes();
        var options = PdfPowerPointImportOptions.CreateHybrid();
        options.MaxRowsPerSlide = 2;
        options.MaxColumnsPerSlide = 2;

        PdfPowerPointConversionResult result = PdfCore.PdfDocument.Load(pdf).ToPowerPointPresentationResult(options);

        Assert.Equal(4, result.Report.TableEntries.Count);
        Assert.Equal(4, result.Report.VisualPages.Count);
        Assert.All(result.Report.TableEntries, entry => {
            Assert.Equal(4, entry.SegmentCount);
            Assert.InRange(entry.RowCount, 1, 2);
            Assert.InRange(entry.ColumnCount, 1, 2);
            Assert.True(entry.HeaderRowIncluded);
        });

        using var presentation = new MemoryStream();
        using (result.Value) result.Value.Save(presentation);
        using PresentationDocument package = PresentationDocument.Open(new MemoryStream(presentation.ToArray()), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
        Assert.Equal(4, package.PresentationPart!.SlideParts.Count());
        Assert.All(package.PresentationPart.SlideParts, slidePart => {
            Assert.Single(slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Picture>());
            List<A.Table> slideTables = slidePart.Slide.Descendants<A.Table>().ToList();
            Assert.InRange(slideTables.Count, 1, 2);
            Assert.All(slideTables, table =>
                Assert.InRange(table.TableGrid!.Elements<A.GridColumn>().Count(), 1, 2));
            Assert.True(
                slideTables.Count == 1 && slideTables[0].Elements<A.TableRow>().Count() == 3 ||
                slideTables.Count == 2 && slideTables
                    .Select(static table => table.Elements<A.TableRow>().Count())
                    .OrderBy(static count => count)
                    .SequenceEqual(new[] { 1, 2 }));
        });
        List<A.Table> hybridTables = package.PresentationPart.SlideParts
            .SelectMany(part => part.Slide.Descendants<A.Table>())
            .ToList();
        Assert.Equal(6, hybridTables.Count);
        Assert.Contains(hybridTables, table => ContainsRows(table, new[] { "C1", "C2" }));
        Assert.Contains(hybridTables, table => ContainsRows(table, new[] { "C3", "C4" }));
        Assert.Contains(hybridTables, table => ContainsRows(table, new[] { "R3C1", "R3C2" }, new[] { "R4C1", "R4C2" }));
        Assert.Contains(hybridTables, table => ContainsRows(table, new[] { "R3C3", "R3C4" }, new[] { "R4C3", "R4C4" }));
    }

    [Fact]
    public void PdfDocument_ToPowerPointPresentation_HybridChargesRepeatedBackgroundsToAggregateBudget() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 520,
                PageHeight = 420,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 9
            })
            .Table(new[] {
                new[] { "C1", "C2" },
                new[] { "R1C1", "R1C2" },
                new[] { "R2C1", "R2C2" },
                new[] { "R3C1", "R3C2" }
            }, style: new PdfCore.PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 120, 120 },
                HeaderRowCount = 1
            })
            .ToBytes();
        PdfCore.PdfDocument source = PdfCore.PdfDocument.Load(pdf);
        PdfCore.PdfPageRenderResult rendered = Assert.Single(source.Render.Pages(
            options: new PdfCore.PdfPageRenderOptions { Dpi = 144 }));
        long oneBackgroundBytes = Assert.IsType<byte[]>(rendered.Bytes).LongLength;
        var options = PdfPowerPointImportOptions.CreateHybrid();
        options.MaxRowsPerSlide = 1;
        options.MaxTotalOutputBytes = oneBackgroundBytes * 2;

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            source.ToPowerPointPresentationResult(options));

        Assert.Contains("aggregate output byte limit", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfDocument_ToPowerPointPresentation_HybridUsesVisualPlacementForMixedPageAspectRatios() {
        byte[] source = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 420,
                PageHeight = 240,
                MarginLeft = 24,
                MarginRight = 24,
                MarginTop = 24,
                MarginBottom = 24,
                DefaultFontSize = 9
            })
            .Paragraph(p => p.Text("Landscape reference"))
            .PageBreak()
            .Table(new[] {
                new[] { "Code", "Qty" },
                new[] { "A-100", "2" },
                new[] { "B-200", "14" }
            }, style: new PdfCore.PdfTableStyle {
                HeaderRowCount = 1,
                ColumnWidthPoints = new List<double?> { 100, 60 },
                CellPaddingX = 4,
                CellPaddingY = 3
            })
            .ToBytes();
        PdfCore.PdfDocument resized = PdfCore.PdfDocument.Load(source).Pages.SetMediaBox(0, 0, 240, 420, 2);

        PdfPowerPointConversionResult result = resized.ToPowerPointPresentationResult(PdfPowerPointImportOptions.CreateHybrid());

        using var presentation = new MemoryStream();
        using (result.Value) result.Value.Save(presentation);
        using PresentationDocument package = PresentationDocument.Open(new MemoryStream(presentation.ToArray()), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
        SlidePart portraitSlide = package.PresentationPart!.SlideParts.ElementAt(1);
        DocumentFormat.OpenXml.Presentation.Picture picture = Assert.Single(
            portraitSlide.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Picture>());
        DocumentFormat.OpenXml.Presentation.GraphicFrame frame = Assert.Single(
            portraitSlide.Slide.Descendants<DocumentFormat.OpenXml.Presentation.GraphicFrame>());
        A.Transform2D pictureTransform = picture.ShapeProperties!.Transform2D!;
        DocumentFormat.OpenXml.Presentation.Transform frameTransform = frame.Transform!;
        long pictureLeft = pictureTransform.Offset!.X!.Value;
        long pictureTop = pictureTransform.Offset.Y!.Value;
        long pictureRight = pictureLeft + pictureTransform.Extents!.Cx!.Value;
        long pictureBottom = pictureTop + pictureTransform.Extents.Cy!.Value;
        long tableLeft = frameTransform.Offset!.X!.Value;
        long tableTop = frameTransform.Offset.Y!.Value;
        long tableRight = tableLeft + frameTransform.Extents!.Cx!.Value;
        long tableBottom = tableTop + frameTransform.Extents.Cy!.Value;

        Assert.True(pictureLeft > 0);
        Assert.InRange(tableLeft, pictureLeft, pictureRight);
        Assert.InRange(tableRight, pictureLeft, pictureRight);
        Assert.InRange(tableTop, pictureTop, pictureBottom);
        Assert.InRange(tableBottom, pictureTop, pictureBottom);
    }

    [Fact]
    public void PdfDocument_ToPowerPointPresentation_HybridRetainsEditableTablesWhenVisualRenderFails() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Table(new[] {
                new[] { "Code", "Qty" },
                new[] { "A-100", "2" },
                new[] { "B-200", "14" }
            }, style: new PdfCore.PdfTableStyle {
                HeaderRowCount = 1,
                ColumnWidthPoints = new List<double?> { 180, 80 }
            })
            .ToBytes();
        var options = PdfPowerPointImportOptions.CreateHybrid();
        options.MaxPixelsPerPage = 10;

        PdfPowerPointConversionResult result = PdfCore.PdfDocument.Load(pdf).ToPowerPointPresentationResult(options);

        Assert.False(Assert.Single(result.Report.VisualPages).Succeeded);
        Assert.Single(result.Report.TableEntries);
        Assert.True(result.Report.HasOmittedPageContent);
        Assert.Contains(
            result.Warnings,
            warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Warning &&
                       warning.Details.TryGetValue("Disposition", out string? disposition) &&
                       disposition == "Omitted");
        using var presentation = new MemoryStream();
        using (result.Value) result.Value.Save(presentation);
        using PresentationDocument package = PresentationDocument.Open(new MemoryStream(presentation.ToArray()), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
        SlidePart slide = Assert.Single(package.PresentationPart!.SlideParts);
        Assert.Empty(slide.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Picture>());
        Assert.Single(slide.Slide.Descendants<A.Table>());
    }

    [Fact]
    public void PdfPowerPointConversionReport_HybridCorrelatesRenderFailuresWithSourcePages() {
        var successfulRender = new PdfCore.PdfPageRenderResult(
            1,
            PdfCore.PdfPageRenderFormat.Png,
            new byte[] { 1 },
            1,
            1,
            TimeSpan.Zero,
            Array.Empty<PdfCore.PdfRenderCapabilityDiagnostic>());
        var failedRender = new PdfCore.PdfPageRenderResult(
            2,
            PdfCore.PdfPageRenderFormat.Png,
            null,
            0,
            0,
            TimeSpan.Zero,
            Array.Empty<PdfCore.PdfRenderCapabilityDiagnostic>(),
            new[] { "render failed" });
        var sourceScope = new PdfCore.PdfTableExtractionScopeReport(
            sourcePageCount: 2,
            pagesWithTables: 1,
            detectedTableCount: 1,
            nonTableTextBlockCount: 0,
            vectorPrimitiveCount: 0,
            imageCount: 1,
            linkCount: 0,
            formWidgetCount: 0,
            annotationCount: 0,
            pageActionCount: 0,
            optionalContentGroupCount: 0,
            interactiveMediaAnnotationCount: 0,
            analysisTruncated: false);
        var failedScope = new PdfCore.PdfTableExtractionScopeReport(
            sourcePageCount: 1,
            pagesWithTables: 1,
            detectedTableCount: 1,
            nonTableTextBlockCount: 0,
            vectorPrimitiveCount: 0,
            imageCount: 0,
            linkCount: 0,
            formWidgetCount: 0,
            annotationCount: 0,
            pageActionCount: 0,
            optionalContentGroupCount: 0,
            interactiveMediaAnnotationCount: 0,
            analysisTruncated: false);
        var report = new PdfPowerPointConversionReport(
            Array.Empty<PdfPowerPointTableImportEntry>(),
            new[] {
                new PdfPowerPointVisualPageEntry(successfulRender, slideIndex: 0),
                new PdfPowerPointVisualPageEntry(failedRender, slideIndex: 1)
            },
            sourceScope,
            failedScope);

        Assert.False(report.HasOmittedPageContent);
        PdfCore.PdfConversionWarning imageWarning = Assert.Single(
            report.Warnings,
            static warning => warning.Code == "PdfImagesNotEditable");
        Assert.Equal(PdfCore.PdfConversionWarningSeverity.Information, imageWarning.Severity);
        Assert.Equal("VisualOnly", imageWarning.Details["Disposition"]);
    }

    [Fact]
    public void PdfTables_SaveTablesAsPowerPoint_ImportsDetectedTablesAsPowerPointTables() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "A-100", "Alpha", "2" },
                new[] { "B-200", "Beta", "14" }
            }, style: new PdfCore.PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 70, 170, 60 },
                HeaderRowCount = 1,
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .ToBytes();

        using var presentation = new MemoryStream();
        PdfPowerPointConversionReport report = PowerPointPdfConverterExtensions.SaveAsPowerPoint(
            LoadTables(pdf),
            presentation,
            PdfPowerPointImportOptions.CreateEditableTables());

        PdfPowerPointTableImportEntry result = Assert.Single(report.TableEntries);
        Assert.Equal(1, result.PageNumber);
        Assert.Equal(0, result.TableIndex);
        Assert.Equal(0, result.SlideIndex);
        Assert.Equal(3, result.ColumnCount);
        Assert.Equal(2, result.RowCount);
        Assert.False(result.Truncated);
        Assert.True(result.HeaderRowIncluded);

        using PresentationDocument package = PresentationDocument.Open(new MemoryStream(presentation.ToArray()), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());

        A.Table table = GetSingleTable(package);
        Assert.True(table.TableProperties?.FirstRow?.Value ?? false);
        Assert.True(table.TableProperties?.BandRow?.Value ?? false);

        List<A.TableRow> rows = table.Elements<A.TableRow>().ToList();
        Assert.Equal(3, rows.Count);
        Assert.Equal(new[] { "Code", "Name", "Qty" }, ReadRowText(rows[0]));
        Assert.Equal(new[] { "A-100", "Alpha", "2" }, ReadRowText(rows[1]));
        Assert.Equal(new[] { "B-200", "Beta", "14" }, ReadRowText(rows[2]));
        Assert.Null(ReadHorizontalAlignment(rows[0], 2));
        Assert.Null(ReadHorizontalAlignment(rows[1], 1));
        Assert.Equal(A.TextAlignmentTypeValues.Right, ReadHorizontalAlignment(rows[1], 2));
        Assert.Equal(A.TextAlignmentTypeValues.Right, ReadHorizontalAlignment(rows[2], 2));
        long[] columnWidths = ReadColumnWidths(table);
        Assert.Equal(3, columnWidths.Length);
        Assert.True(columnWidths[1] > columnWidths[0]);
        Assert.True(columnWidths[1] > columnWidths[2]);
        Assert.Contains(ReadAllText(package), text => text == "PDF page 1, table 1");
    }

    [Fact]
    public void PdfTables_SaveTablesAsPowerPoint_AppliesRowCapsAndKeepsPresentationValidWhenEmpty() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .KeyValueTable(new[] {
                PdfCore.PdfKeyValueRow.Text("InvoiceId", "INV-001"),
                PdfCore.PdfKeyValueRow.Text("Customer", "Evotec"),
                PdfCore.PdfKeyValueRow.Text("Due", "2026-06-30")
            }, style: new PdfCore.PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 120, 170 },
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .PageBreak()
            .Paragraph(p => p.Text("No table on this page."))
            .ToBytes();

        using var presentation = new MemoryStream();
        PdfPowerPointConversionReport report = PowerPointPdfConverterExtensions.SaveAsPowerPoint(
            LoadTables(pdf, PdfCore.PdfPageRange.From(1, 1)),
            presentation,
            new PdfPowerPointImportOptions {
                Mode = PdfPowerPointImportMode.EditableTables,
                MaxRows = 2,
                IncludeSourceTitles = false
            });

        PdfPowerPointTableImportEntry result = Assert.Single(report.TableEntries);
        Assert.Equal(2, result.RowCount);
        Assert.Equal(3, result.TotalRowCount);
        Assert.True(result.Truncated);
        Assert.True(result.HeaderRowIncluded);
        Assert.True(report.HasLoss);
        Assert.Throws<InvalidOperationException>(() => report.RequireNoLoss());

        using (PresentationDocument package = PresentationDocument.Open(new MemoryStream(presentation.ToArray()), false)) {
            Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
            A.Table table = GetSingleTable(package);
            List<A.TableRow> rows = table.Elements<A.TableRow>().ToList();
            Assert.Equal(3, rows.Count);
            Assert.Equal(new[] { "Key", "Value" }, ReadRowText(rows[0]));
            Assert.Equal(new[] { "InvoiceId", "INV-001" }, ReadRowText(rows[1]));
            Assert.Equal(new[] { "Customer", "Evotec" }, ReadRowText(rows[2]));
        }

        using var emptyPresentation = new MemoryStream();
        PdfPowerPointConversionReport emptyReport = PowerPointPdfConverterExtensions.SaveAsPowerPoint(
            LoadTables(pdf, PdfCore.PdfPageRange.From(2, 2)),
            emptyPresentation,
            new PdfPowerPointImportOptions {
                Mode = PdfPowerPointImportMode.EditableTables,
                EmptyPresentationMessage = "Nothing tabular was detected."
            });

        Assert.Empty(emptyReport.TableEntries);
        using PresentationDocument emptyPackage = PresentationDocument.Open(new MemoryStream(emptyPresentation.ToArray()), false);
        Assert.Empty(new OpenXmlValidator().Validate(emptyPackage).ToList());
        Assert.Empty(emptyPackage.PresentationPart!.SlideParts.SelectMany(part => part.Slide.Descendants<A.Table>()));
        Assert.Contains(ReadAllText(emptyPackage), text => text == "Nothing tabular was detected.");
    }

    [Fact]
    public void PdfTables_SaveTablesAsPowerPoint_SkipsHeaderOnlySegmentsWhenHeadersAreDisabled() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 420,
                PageHeight = 260,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Table(new[] {
                new[] { "Code", "Name" }
            }, style: new PdfCore.PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 90, 150 },
                HeaderRowCount = 1,
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .ToBytes();

        using var presentation = new MemoryStream();
        PdfPowerPointConversionReport report = PowerPointPdfConverterExtensions.SaveAsPowerPoint(
            LoadTables(pdf),
            presentation,
            new PdfPowerPointImportOptions {
                Mode = PdfPowerPointImportMode.EditableTables,
                IncludeColumnHeaderRows = false,
                EmptyPresentationMessage = "No table rows were imported."
            });

        Assert.Empty(report.TableEntries);
        using PresentationDocument package = PresentationDocument.Open(new MemoryStream(presentation.ToArray()), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
        Assert.Empty(package.PresentationPart!.SlideParts.SelectMany(part => part.Slide.Descendants<A.Table>()));
        Assert.Contains(ReadAllText(package), text => text == "No table rows were imported.");
    }

    [Fact]
    public void PdfTables_SaveTablesAsPowerPoint_SplitsLargeTablesAcrossSlides() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 520,
                PageHeight = 420,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 9
            })
            .Table(new[] {
                new[] { "C1", "C2", "C3", "C4" },
                new[] { "R1C1", "R1C2", "R1C3", "R1C4" },
                new[] { "R2C1", "R2C2", "R2C3", "R2C4" },
                new[] { "R3C1", "R3C2", "R3C3", "R3C4" },
                new[] { "R4C1", "R4C2", "R4C3", "R4C4" }
            }, style: new PdfCore.PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 80, 80, 80, 80 },
                HeaderRowCount = 1,
                CellPaddingX = 4,
                CellPaddingY = 3
            })
            .ToBytes();

        using var presentation = new MemoryStream();
        PdfPowerPointConversionReport report = PowerPointPdfConverterExtensions.SaveAsPowerPoint(
            LoadTables(pdf),
            presentation,
            new PdfPowerPointImportOptions {
                Mode = PdfPowerPointImportMode.EditableTables,
                MaxRowsPerSlide = 2,
                MaxColumnsPerSlide = 2
            });

        IReadOnlyList<PdfPowerPointTableImportEntry> results = report.TableEntries;
        Assert.Equal(4, results.Count);
        Assert.All(results, result => {
            Assert.Equal(4, result.SegmentCount);
            Assert.Equal(4, result.SourceColumnCount);
            Assert.Equal(2, result.ColumnCount);
            Assert.Equal(2, result.RowCount);
            Assert.Equal(4, result.TotalRowCount);
            Assert.True(result.HeaderRowIncluded);
        });
        Assert.Equal(new[] { 0, 1, 2, 3 }, results.Select(result => result.SegmentIndex).ToArray());
        Assert.Equal(new[] { 0, 0, 2, 2 }, results.Select(result => result.RowStartIndex).ToArray());
        Assert.Equal(new[] { 0, 2, 0, 2 }, results.Select(result => result.ColumnStartIndex).ToArray());
        Assert.Equal(new[] { 0, 1, 2, 3 }, results.Select(result => result.SlideIndex).ToArray());

        using PresentationDocument package = PresentationDocument.Open(new MemoryStream(presentation.ToArray()), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());

        List<A.Table> tables = package.PresentationPart!.SlideParts
            .SelectMany(part => part.Slide.Descendants<A.Table>())
            .ToList();
        Assert.Equal(4, tables.Count);
        Assert.Contains(tables, table => ContainsRows(table, new[] { "C1", "C2" }, new[] { "R1C1", "R1C2" }, new[] { "R2C1", "R2C2" }));
        Assert.Contains(tables, table => ContainsRows(table, new[] { "C3", "C4" }, new[] { "R1C3", "R1C4" }, new[] { "R2C3", "R2C4" }));
        Assert.Contains(tables, table => ContainsRows(table, new[] { "C1", "C2" }, new[] { "R3C1", "R3C2" }, new[] { "R4C1", "R4C2" }));
        Assert.Contains(tables, table => ContainsRows(table, new[] { "C3", "C4" }, new[] { "R3C3", "R3C4" }, new[] { "R4C3", "R4C4" }));

        string[] text = ReadAllText(package);
        Assert.Contains(text, value => value == "PDF page 1, table 1 (part 1 of 4)");
        Assert.Contains(text, value => value == "PDF page 1, table 1 (part 4 of 4)");
    }

    [Fact]
    public void PdfTables_ToPowerPoint_MergesPageContinuationsAndRepeatedHeaders() {
        var rows = new List<string[]> {
            new[] { "Group", "State" },
            new[] { "Metric", "Owner" }
        };
        for (int index = 1; index <= 30; index++) {
            rows.Add(new[] { "Check " + index, "Team " + index });
        }
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFontSize = 9
            })
            .Table(rows, style: new PdfCore.PdfTableStyle {
                HeaderRowCount = 2,
                RepeatHeaderRowCount = 2,
                ColumnWidthPoints = new List<double?> { 120, 120 },
                CellPaddingX = 5,
                CellPaddingY = 3
            })
            .ToBytes();

        PdfPowerPointConversionResult result = LoadTables(pdf).ToPowerPointPresentationResult(new PdfPowerPointImportOptions {
            Mode = PdfPowerPointImportMode.EditableTables,
            SuppressRepeatedBodyHeaderRows = true
        });
        PdfPowerPointTableImportEntry entry = Assert.Single(result.Report.TableEntries);

        Assert.True(entry.SourceTableCount > 1);
        Assert.Equal(entry.SourceTableCount, entry.SourcePageNumbers.Count);
        Assert.Equal(Enumerable.Range(1, entry.SourceTableCount), entry.SourcePageNumbers);
        Assert.Equal(30, entry.RowCount);
        Assert.Equal(30, entry.TotalRowCount);
        Assert.Equal(1, entry.AdditionalHeaderRowCount);
        Assert.Equal(entry.SourceTableCount - 1, entry.SuppressedRepeatedHeaderRows);

        using var presentation = new MemoryStream();
        using (result.Value) result.Value.Save(presentation);
        using PresentationDocument package = PresentationDocument.Open(new MemoryStream(presentation.ToArray()), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
        A.Table table = GetSingleTable(package);
        string[][] tableRows = table.Elements<A.TableRow>().Select(ReadRowText).ToArray();
        Assert.Equal(new[] { "Group / Metric", "State / Owner" }, tableRows[0]);
        Assert.Equal(new[] { "Check 1", "Team 1" }, tableRows[1]);
        Assert.Equal(new[] { "Check 30", "Team 30" }, tableRows[30]);
    }

    private static PdfCore.PdfDocumentReadResult LoadTables(byte[] pdf, params PdfCore.PdfPageRange[] ranges) {
        var layout = new PdfCore.PdfTextLayoutOptions { ForceSingleColumn = true };
        return ranges.Length == 0
            ? PdfCore.PdfDocumentReadResult.Load(pdf, layout)
            : PdfCore.PdfDocumentReadResult.LoadPageRanges(pdf, layout, ranges);
    }

    private static A.Table GetSingleTable(PresentationDocument package) {
        return Assert.Single(package.PresentationPart!.SlideParts.SelectMany(part => part.Slide.Descendants<A.Table>()));
    }

    private static bool ContainsRows(A.Table table, params string[][] expectedRows) {
        string[][] rows = table.Elements<A.TableRow>()
            .Select(ReadRowText)
            .ToArray();
        if (rows.Length != expectedRows.Length) {
            return false;
        }

        for (int rowIndex = 0; rowIndex < expectedRows.Length; rowIndex++) {
            if (!rows[rowIndex].SequenceEqual(expectedRows[rowIndex])) {
                return false;
            }
        }

        return true;
    }

    private static long[] ReadColumnWidths(A.Table table) {
        return table.TableGrid!.Elements<A.GridColumn>()
            .Select(column => column.Width?.Value ?? 0L)
            .ToArray();
    }

    private static A.TextAlignmentTypeValues? ReadHorizontalAlignment(A.TableRow row, int columnIndex) {
        return row.Elements<A.TableCell>()
            .ElementAt(columnIndex)
            .TextBody?
            .Elements<A.Paragraph>()
            .FirstOrDefault()?
            .ParagraphProperties?
            .Alignment?
            .Value;
    }

    private static string[] ReadRowText(A.TableRow row) {
        return row.Elements<A.TableCell>()
            .Select(cell => string.Concat(cell.Descendants<A.Text>().Select(text => text.Text ?? string.Empty)))
            .ToArray();
    }

    private static string[] ReadAllText(PresentationDocument package) {
        return package.PresentationPart!.SlideParts
            .SelectMany(part => part.Slide.Descendants<A.Text>())
            .Select(text => text.Text ?? string.Empty)
            .ToArray();
    }
}
