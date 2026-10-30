using OfficeIMO.Pdf.PowerPoint;
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
    public void PdfDocument_ToPowerPointDocument_VisualProfileCreatesOnePageImagePerSlide() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(360),
                PageHeight = global::OfficeIMO.OfficeLength.Points(240),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(24),
                MarginRight = global::OfficeIMO.OfficeLength.Points(24),
                MarginTop = global::OfficeIMO.OfficeLength.Points(24),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(24)
            })
            .H1("Visual page one")
            .PageBreak()
            .Paragraph(p => p.Text("Visual page two"))
            .ToBytes();

        PdfCore.PdfDocument opened = PdfCore.PdfDocument.Load(pdf);
        PdfPowerPointConversionResult result = opened.ConvertToPowerPointWithEvidence(
            PowerPointPdfImportOptions.CreateVisualPages());

        Assert.Equal(PowerPointPdfImportMode.VisualPages, result.Report.Mode);
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
    public void PdfDocument_ToPowerPointDocument_DefaultCreatesNativeEditableObjects() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(420),
                PageHeight = global::OfficeIMO.OfficeLength.Points(360),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(36),
                MarginRight = global::OfficeIMO.OfficeLength.Points(36),
                MarginTop = global::OfficeIMO.OfficeLength.Points(36),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(36),
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
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
                ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(160), global::OfficeIMO.OfficeLength.Points(90) }
            })
            .ToBytes();

        Assert.Equal(PowerPointPdfImportMode.Auto, new PowerPointPdfImportOptions().Mode);

        PdfPowerPointConversionResult result = PdfCore.PdfDocument.Load(pdf)
            .ConvertToPowerPointWithEvidence();

        Assert.Equal(PowerPointPdfImportMode.EditableContent, result.Report.Mode);
        PdfPowerPointEditablePageEntry page = Assert.Single(result.Report.EditablePages);
        Assert.True(page.TextBoxCount >= 2);
        Assert.Equal(1, page.TableCount);
        Assert.Equal(1, page.ImageCount);
        Assert.Single(result.Report.TableEntries);
        Assert.Contains(result.Warnings, warning => warning.Code == "pdf-editable-content-reconstructed");

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
    public void PdfDocument_ToPowerPointDocument_EditableContentCreatesNativeSafeVectorShapes() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(420),
                PageHeight = global::OfficeIMO.OfficeLength.Points(300),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(36),
                MarginRight = global::OfficeIMO.OfficeLength.Points(36),
                MarginTop = global::OfficeIMO.OfficeLength.Points(36),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(36),
                PageBackgroundShapes = new[] {
                    PdfCore.PdfPageBackgroundShape.Rectangle(
                        40,
                        210,
                        120,
                        36,
                        fill: global::OfficeIMO.OfficeColor.FromRgb(219, 234, 254),
                        stroke: global::OfficeIMO.OfficeColor.FromRgb(37, 99, 235),
                        strokeWidth: 1)
                }
            })
            .Paragraph(p => p.Text("Editable rectangle"))
            .ToBytes();

        PdfPowerPointConversionResult result = PdfCore.PdfDocument.Load(pdf)
            .ConvertToPowerPointWithEvidence(PowerPointPdfImportOptions.CreateEditableContent());

        PdfPowerPointEditablePageEntry page = Assert.Single(result.Report.EditablePages);
        Assert.True(page.ShapeCount >= 1);
        Assert.Equal(0, page.OmittedVectorCount);
        using var presentation = new MemoryStream();
        using (result.Value) result.Value.Save(presentation);
        using PresentationDocument package = PresentationDocument.Open(new MemoryStream(presentation.ToArray()), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
        SlidePart slide = Assert.Single(package.PresentationPart!.SlideParts);
        Assert.Contains(
            slide.Slide.Descendants<A.PresetGeometry>(),
            geometry => geometry.Preset?.Value == A.ShapeTypeValues.Rectangle);
    }

    [Fact]
    public void LogicalEditableContentReportsRepresentableVectorsAsOmittedWithoutRenderer() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(420),
                PageHeight = global::OfficeIMO.OfficeLength.Points(300),
                PageBackgroundShapes = new[] {
                    PdfCore.PdfPageBackgroundShape.Rectangle(
                        40,
                        210,
                        120,
                        36,
                        fill: global::OfficeIMO.OfficeColor.FromRgb(219, 234, 254),
                        stroke: global::OfficeIMO.OfficeColor.FromRgb(37, 99, 235),
                        strokeWidth: 1)
                }
            })
            .Paragraph(paragraph => paragraph.Text("Logical-only vector"))
            .ToBytes();
        PdfCore.PdfDocumentReadResult logical = PdfCore.PdfDocumentReadResult.Load(pdf);
        PdfCore.PdfLogicalPage sourcePage = Assert.Single(logical.Pages);
        Assert.True(sourcePage.VectorPrimitiveCount > 0);

        PdfPowerPointConversionResult result = logical.ConvertToPowerPointWithEvidence(
            PowerPointPdfImportOptions.CreateEditableContent());

        using (result.Value) {
            PdfPowerPointEditablePageEntry page = Assert.Single(result.Report.EditablePages);
            Assert.Equal(0, page.ShapeCount);
            Assert.Equal(sourcePage.VectorPrimitiveCount, page.OmittedVectorCount);
            Assert.Contains(result.Warnings, static warning =>
                warning.Code == "pdf-vectors-not-reconstructed" &&
                warning.LossKind == OfficeConversionLossKind.Omission);
        }
    }

    [Theory]
    [InlineData(PowerPointPdfImportMode.VisualPages)]
    [InlineData(PowerPointPdfImportMode.EditableTables)]
    [InlineData(PowerPointPdfImportMode.HybridVisualAndEditableTables)]
    [InlineData(PowerPointPdfImportMode.EditableContent)]
    [InlineData(PowerPointPdfImportMode.Auto)]
    public void PdfDocument_ToPowerPointDocument_EnforcesPageLimitInEveryMode(
        PowerPointPdfImportMode mode) {
        byte[] pdf = PdfCore.PdfDocument.Create()
            .Paragraph(p => p.Text("Page one"))
            .PageBreak()
            .Paragraph(p => p.Text("Page two"))
            .ToBytes();
        var options = new PowerPointPdfImportOptions {
            Mode = mode,
            MaxPages = 1
        };

        Exception exception = Assert.ThrowsAny<Exception>(() =>
            PdfCore.PdfDocument.Load(pdf).ConvertToPowerPointWithEvidence(options));

        Assert.Contains("limit", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(PowerPointPdfImportMode.EditableTables)]
    [InlineData(PowerPointPdfImportMode.EditableContent)]
    public void PdfDocument_ToPowerPointDocument_RejectsOversizedSelectionBeforeLogicalExtraction(
        PowerPointPdfImportMode mode) {
        byte[] pdf = PdfCore.PdfDocument.Create()
            .Paragraph(p => p.Text("Only one source page"))
            .ToBytes();
        var options = new PowerPointPdfImportOptions {
            Mode = mode,
            MaxPages = 1,
            ReadOptions = new PdfCore.PdfLoadOptions {
                PageSelection = PdfCore.PdfPageSelection.From(2, 2)
            }
        };

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            PdfCore.PdfDocument.Load(pdf).ConvertToPowerPointWithEvidence(options));

        Assert.Contains("page count 2", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("configured limit of 1", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void PdfDocument_ToPowerPointDocument_AllowsRaisedDestinationAndSemanticPageLimits() {
        const int selectedPageCount = 1_001;
        byte[] pdf = PdfCore.PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Repeated source page"))
            .ToBytes();
        var options = new PowerPointPdfImportOptions {
            Mode = PowerPointPdfImportMode.EditableTables,
            MaxPages = selectedPageCount,
            ReadOptions = new PdfCore.PdfLoadOptions {
                PageSelection = PdfCore.PdfPageSelection.From(Enumerable.Repeat(1, selectedPageCount).ToArray()),
                Pipeline = new PdfCore.PdfUnderstandingPipelineOptions { MaxPages = selectedPageCount }
            }
        };

        PdfPowerPointConversionResult result = PdfCore.PdfDocument.Load(pdf)
            .ConvertToPowerPointWithEvidence(options);

        using (result.Value) {
            Assert.Single(result.Value.Slides);
        }
    }

    [Fact]
    public void PdfDocumentReadResult_ToPowerPointDocument_AppliesSelectionBeforeDestinationLimit() {
        byte[] pdf = PdfCore.PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Page one"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Page two"))
            .ToBytes();
        PdfCore.PdfDocumentReadResult logical = PdfCore.PdfDocument.Load(pdf).Read();
        var options = PowerPointPdfImportOptions.CreateEditableContent();
        options.MaxPages = 1;
        options.ReadOptions = new PdfCore.PdfLoadOptions {
            PageSelection = PdfCore.PdfPageSelection.From(2)
        };

        PdfPowerPointConversionResult result = logical.ConvertToPowerPointWithEvidence(options);

        using (result.Value) {
            PdfPowerPointEditablePageEntry page = Assert.Single(result.Report.EditablePages);
            Assert.Equal(2, page.PageNumber);
        }
    }

    [Fact]
    public void PdfDocumentReadResult_ToPowerPointDocument_PreservesSelectedOrderAndDuplicates() {
        byte[] pdf = PdfCore.PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Page one"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Page two"))
            .ToBytes();
        PdfCore.PdfDocumentReadResult logical = PdfCore.PdfDocument.Load(pdf).Read();
        var options = PowerPointPdfImportOptions.CreateEditableContent();
        options.MaxPages = 3;
        options.ReadOptions = new PdfCore.PdfLoadOptions {
            PageSelection = PdfCore.PdfPageSelection.From(2, 1, 2)
        };

        PdfPowerPointConversionResult result = logical.ConvertToPowerPointWithEvidence(options);

        using (result.Value) {
            Assert.Equal(new[] { 2, 1, 2 }, result.Report.EditablePages.Select(page => page.PageNumber).ToArray());
        }
    }

    [Fact]
    public void PdfDocument_ToPowerPointDocument_EditableObjectLimitReportsTextAndTableLoss() {
        byte[] pdf = PdfCore.PdfDocument.Create()
            .H1("Limit contract")
            .Paragraph(p => p.Text("Second editable text block"))
            .Table(new[] {
                new[] { "Metric", "Value" },
                new[] { "Ready", "Yes" },
                new[] { "Validated", "Yes" }
            }, style: new PdfCore.PdfTableStyle { HeaderRowCount = 1 })
            .ToBytes();
        var options = PowerPointPdfImportOptions.CreateEditableContent();
        options.MaxEditableObjectsPerPage = 1;

        PdfPowerPointConversionResult result = PdfCore.PdfDocument.Load(pdf)
            .ConvertToPowerPointWithEvidence(options);

        using (result.Value) {
            PdfPowerPointEditablePageEntry page = Assert.Single(result.Report.EditablePages);
            Assert.True(page.OmittedTextCount > 0);
            Assert.True(page.OmittedTableCount > 0);
            Assert.True(page.HasOmittedContent);
            Assert.True(result.HasLoss);
            Assert.Contains(result.Warnings, warning => warning.Code == "pdf-editable-object-limit-reached");
            Assert.Throws<InvalidOperationException>(result.Report.RequireNoLoss);
        }
    }

    [Theory]
    [InlineData(360D, 180D, 2000)]
    [InlineData(1440D, 720D, 500)]
    public void PdfDocument_ToPowerPointDocument_EditableTextScalesWithPageGeometry(
        double pageWidth,
        double pageHeight,
        int expectedFontSizeHundredths) {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(pageWidth),
                PageHeight = global::OfficeIMO.OfficeLength.Points(pageHeight),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(12),
                MarginRight = global::OfficeIMO.OfficeLength.Points(12),
                MarginTop = global::OfficeIMO.OfficeLength.Points(12),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(12),
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
            })
            .Paragraph(p => p.Text("Scale probe"))
            .ToBytes();
        PdfPowerPointConversionResult result = PdfCore.PdfDocument.Load(pdf)
            .ConvertToPowerPointWithEvidence(PowerPointPdfImportOptions.CreateEditableContent());

        using var presentation = new MemoryStream();
        using (result.Value) result.Value.Save(presentation);
        using PresentationDocument package = PresentationDocument.Open(new MemoryStream(presentation.ToArray()), false);
        A.Run run = Assert.Single(
            package.PresentationPart!.SlideParts.Single().Slide.Descendants<A.Run>(),
            candidate => candidate.Text?.Text == "Scale probe");
        Assert.Equal(expectedFontSizeHundredths, run.RunProperties?.FontSize?.Value);
    }

    [Theory]
    [InlineData(0, 0, 0)]
    [InlineData(0, 90, 90)]
    [InlineData(0, 180, 180)]
    [InlineData(0, 270, 270)]
    [InlineData(90, 0, 90)]
    [InlineData(90, 90, 180)]
    [InlineData(90, 180, 270)]
    [InlineData(90, 270, 0)]
    public void PdfDocument_ToPowerPointDocument_EditableTextUsesVisualPageRotation(
        int sourceTextRotation,
        int pageRotation,
        int expectedPowerPointRotation) {
        byte[] source = PdfCore.PdfDocument.Create(new PdfCore.PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(360),
                PageHeight = global::OfficeIMO.OfficeLength.Points(240),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(24),
                MarginRight = global::OfficeIMO.OfficeLength.Points(24),
                MarginTop = global::OfficeIMO.OfficeLength.Points(24),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(24)
            })
            .Paragraph(p => p.Text("Rotation probe"))
            .ToBytes();
        if (sourceTextRotation != 0) {
            source = PdfCore.PdfDocument.Load(source).Pages.Rotate(sourceTextRotation)
                .Pages.NormalizeRotation().ToBytes();
        }
        byte[] rotated = PdfCore.PdfDocument.Load(source).Pages.Rotate(pageRotation, "1").ToBytes();
        PdfPowerPointConversionResult result = PdfCore.PdfDocument.Load(rotated)
            .ConvertToPowerPointWithEvidence(PowerPointPdfImportOptions.CreateEditableContent());

        using var presentation = new MemoryStream();
        using (result.Value) result.Value.Save(presentation);
        using PresentationDocument package = PresentationDocument.Open(new MemoryStream(presentation.ToArray()), false);
        P.Shape shape = Assert.Single(
            package.PresentationPart!.SlideParts.Single().Slide.Descendants<P.Shape>(),
            candidate => candidate.TextBody?.InnerText == "Rotation probe");
        A.Transform2D transform = shape.ShapeProperties!.Transform2D!;
        Assert.Equal(expectedPowerPointRotation * 60000, transform.Rotation?.Value ?? 0);
        Assert.True(transform.Extents!.Cx!.Value > transform.Extents.Cy!.Value);
    }

    [Fact]
    public void PdfDocument_ToPowerPointDocument_EditableContinuationTablesStayInsideSlide() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(3600D),
                PageHeight = global::OfficeIMO.OfficeLength.Points(1800D),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(36D),
                MarginRight = global::OfficeIMO.OfficeLength.Points(36D),
                MarginTop = global::OfficeIMO.OfficeLength.Points(36D),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(36D),
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10D)
            })
            .Table(new[] {
                new[] { "Metric", "Value" },
                new[] { "One", "1" },
                new[] { "Two", "2" },
                new[] { "Three", "3" }
            }, style: new PdfCore.PdfTableStyle { HeaderRowCount = 1 })
            .ToBytes();
        var options = PowerPointPdfImportOptions.CreateEditableContent();
        options.MaxRowsPerSlide = 1;
        options.IncludeSourceTitles = false;

        PdfPowerPointConversionResult result = PdfCore.PdfDocument.Load(pdf)
            .ConvertToPowerPointWithEvidence(options);

        using (result.Value) {
            Assert.True(result.Value.Slides.Count > 1);
            double slideWidth = result.Value.SlideSize.Width.ToPoints();
            double slideHeight = result.Value.SlideSize.Height.ToPoints();
            double primaryFontSize = Assert.Single(result.Value.Slides[0].Tables).GetCell(0, 0).FontSize?.ToPoints() ?? 0;
            Assert.InRange(primaryFontSize, 1, 3);
            foreach (OfficeIMO.PowerPoint.PowerPointSlide slide in result.Value.Slides.Skip(1)) {
                OfficeIMO.PowerPoint.PowerPointTable table = Assert.Single(slide.Tables);
                Assert.InRange(table.Left.ToPoints(), 0D, slideWidth);
                Assert.InRange(table.Top.ToPoints(), 0D, slideHeight);
                Assert.True(table.Left.ToPoints() + table.Width.ToPoints() <= slideWidth + 0.01D);
                Assert.True(table.Top.ToPoints() + table.Height.ToPoints() <= slideHeight + 0.01D);
                Assert.InRange(table.GetCell(0, 0).FontSize?.ToPoints() ?? 0, 10, 18);
            }
        }
    }

    [Fact]
    public void PdfDocument_ToPowerPointDocument_HybridKeepsVisualPageAndEditableTableOverlay() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(420),
                PageHeight = global::OfficeIMO.OfficeLength.Points(360),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(36),
                MarginRight = global::OfficeIMO.OfficeLength.Points(36),
                MarginTop = global::OfficeIMO.OfficeLength.Points(36),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(36),
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
            })
            .Paragraph(p => p.Text("Quarterly status"))
            .Table(new[] {
                new[] { "Code", "Qty" },
                new[] { "A-100", "2" },
                new[] { "B-200", "14" }
            }, style: new PdfCore.PdfTableStyle {
                HeaderRowCount = 1,
                ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(180), global::OfficeIMO.OfficeLength.Points(80) },
                CellPaddingX = global::OfficeIMO.OfficeLength.Points(6),
                CellPaddingY = global::OfficeIMO.OfficeLength.Points(4)
            })
            .ToBytes();

        PdfPowerPointConversionResult result = PdfCore.PdfDocument.Load(pdf)
            .ConvertToPowerPointWithEvidence(PowerPointPdfImportOptions.CreateHybrid());

        Assert.Equal(PowerPointPdfImportMode.HybridVisualAndEditableTables, result.Report.Mode);
        Assert.Single(result.Report.VisualPages);
        Assert.Single(result.Report.TableEntries);
        Assert.True(result.Report.VisualPages[0].Succeeded);
        Assert.True(result.Report.HasNonEditablePageContent);
        Assert.False(result.Report.HasOmittedPageContent);
        PdfCore.PdfConversionWarning textWarning = Assert.Single(result.Warnings, warning => warning.Code == "pdf-text-not-editable");
        Assert.Equal(PdfCore.PdfConversionWarningSeverity.Information, textWarning.Severity);
        Assert.Equal("VisualOnly", textWarning.Details["Disposition"]);
        Assert.Contains(result.Warnings, warning => warning.Code == "pdf-vectors-not-editable");

        using var presentation = new MemoryStream();
        using (result.Value) result.Value.Save(presentation);
        using PresentationDocument package = PresentationDocument.Open(new MemoryStream(presentation.ToArray()), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
        SlidePart slide = Assert.Single(package.PresentationPart!.SlideParts);
        Assert.Single(slide.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Picture>());
        Assert.Single(slide.Slide.Descendants<A.Table>());
    }

    [Fact]
    public void PdfDocument_ToPowerPointDocument_ReportsDocumentActionsAsOmittedFromHybridAndEditableOutput() {
        var pdfOptions = new PdfCore.PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(420),
            PageHeight = global::OfficeIMO.OfficeLength.Points(360),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(36),
            MarginRight = global::OfficeIMO.OfficeLength.Points(36),
            MarginTop = global::OfficeIMO.OfficeLength.Points(36),
            MarginBottom = global::OfficeIMO.OfficeLength.Points(36),
            DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
        }.SetOpenAction(1, destinationMode: PdfCore.PdfOpenActionDestinationMode.Fit);
        byte[] tablePdf = PdfCore.PdfDocument.Create(pdfOptions)
            .Table(new[] {
                new[] { "Code", "Qty" },
                new[] { "A-100", "2" },
                new[] { "B-200", "14" }
            })
            .ToBytes();
        byte[] pdf = PdfCore.PdfDocument.Load(tablePdf)
            .JavaScript.AddOrReplace("Initialize", "app.alert('open');")
            .ToBytes();

        PdfPowerPointConversionResult hybrid = PdfCore.PdfDocument.Load(pdf)
            .ConvertToPowerPointWithEvidence(PowerPointPdfImportOptions.CreateHybrid());
        using (hybrid.Value) {
            Assert.True(Assert.Single(hybrid.Report.VisualPages).Succeeded);
            Assert.Equal(2, hybrid.Report.SourceScope!.DocumentActionCount);
            Assert.True(hybrid.Report.HasOmittedPageContent);
            Assert.True(hybrid.Report.HasLoss);
            PdfCore.PdfConversionWarning warning = Assert.Single(hybrid.Warnings, static warning =>
                warning.Code == "pdf-document-actions-not-reconstructed");
            Assert.Equal(OfficeConversionLossKind.Omission, warning.LossKind);
            Assert.Equal("2", warning.Details["Count"]);
            Assert.Equal("Omitted", warning.Details["Disposition"]);
            Assert.Throws<InvalidOperationException>(() => hybrid.RequireNoLoss());
        }

        PdfPowerPointConversionResult editable = PdfCore.PdfDocumentReadResult.Load(pdf)
            .ConvertToPowerPointWithEvidence(PowerPointPdfImportOptions.CreateEditableContent());
        using (editable.Value) {
            Assert.Contains(editable.Warnings, static warning =>
                warning.Code == "pdf-document-actions-not-reconstructed" &&
                warning.LossKind == OfficeConversionLossKind.Omission);
            Assert.True(editable.Report.HasLoss);
            Assert.Throws<InvalidOperationException>(() => editable.RequireNoLoss());
        }

        PdfPowerPointConversionResult visual = PdfCore.PdfDocument.Load(pdf)
            .ConvertToPowerPointWithEvidence(PowerPointPdfImportOptions.CreateVisualPages());
        using (visual.Value) {
            Assert.True(visual.Report.HasOmittedPageContent);
            Assert.Contains(visual.Warnings, static warning =>
                warning.Code == "pdf-document-actions-not-reconstructed" &&
                warning.LossKind == OfficeConversionLossKind.Omission);
        }
    }

    [Theory]
    [InlineData(90)]
    [InlineData(270)]
    public void PdfDocument_ToPowerPointDocument_HybridMapsRotatedTableBoundsToVisualCoordinates(int rotation) {
        byte[] source = PdfCore.PdfDocument.Create(new PdfCore.PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(420),
                PageHeight = global::OfficeIMO.OfficeLength.Points(300),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
                MarginRight = global::OfficeIMO.OfficeLength.Points(30),
                MarginTop = global::OfficeIMO.OfficeLength.Points(30),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(30)
            })
            .Table(new[] {
                new[] { "Code", "Qty" },
                new[] { "A-100", "2" },
                new[] { "B-200", "14" }
            }, style: new PdfCore.PdfTableStyle {
                HeaderRowCount = 1,
                ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(140), global::OfficeIMO.OfficeLength.Points(80) }
            })
            .ToBytes();
        byte[] rotated = PdfCore.PdfPageEditor.RotatePages(source, rotation, 1);

        PdfPowerPointConversionResult result = PdfCore.PdfDocument.Load(rotated)
            .ConvertToPowerPointWithEvidence(PowerPointPdfImportOptions.CreateHybrid());
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

        // A clockwise quarter turn puts the source's left-side table near the top; the opposite turn puts it near the bottom.
        if (rotation == 90) Assert.True(tableTop < pictureTop + pictureHeight / 2L);
        else Assert.True(tableTop > pictureTop + pictureHeight / 2L);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
    }

    [Fact]
    public void PdfDocument_ToPowerPointDocument_HybridOmitsSyntheticKeyValueHeader() {
        byte[] pdf = PdfCore.PdfDocument.Create()
            .KeyValueTable(new[] {
                PdfCore.PdfKeyValueRow.Text("InvoiceId", "INV-001"),
                PdfCore.PdfKeyValueRow.Text("Customer", "Evotec"),
                PdfCore.PdfKeyValueRow.Text("Status", "Open")
            }, includeHeader: false)
            .ToBytes();

        PdfPowerPointConversionResult result = PdfCore.PdfDocument.Load(pdf)
            .ConvertToPowerPointWithEvidence(PowerPointPdfImportOptions.CreateHybrid());
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
    public void PdfDocument_ToPowerPointDocument_HybridPreservesSelectedPageIndexesInTableReports() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(420),
                PageHeight = global::OfficeIMO.OfficeLength.Points(300),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
                MarginRight = global::OfficeIMO.OfficeLength.Points(30),
                MarginTop = global::OfficeIMO.OfficeLength.Points(30),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(30)
            })
            .Table(new[] {
                new[] { "Code", "Qty" },
                new[] { "A-100", "2" },
                new[] { "B-200", "14" }
            }, style: new PdfCore.PdfTableStyle { HeaderRowCount = 1, ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(140), global::OfficeIMO.OfficeLength.Points(80) } })
            .PageBreak()
            .Table(new[] {
                new[] { "Name", "Total" },
                new[] { "Alpha", "20" },
                new[] { "Beta", "30" }
            }, style: new PdfCore.PdfTableStyle { HeaderRowCount = 1, ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(140), global::OfficeIMO.OfficeLength.Points(80) } })
            .ToBytes();

        PdfPowerPointConversionResult result = PdfCore.PdfDocument.Load(pdf)
            .ConvertToPowerPointWithEvidence(PowerPointPdfImportOptions.CreateHybrid());

        Assert.Equal(new[] { 0, 1 }, result.Report.TableEntries.Select(entry => entry.PageIndex).ToArray());
        Assert.Equal(new[] { 1, 2 }, result.Report.TableEntries.Select(entry => entry.PageNumber).ToArray());
        result.Value.Dispose();
    }

    [Fact]
    public void PdfDocument_ToPowerPointDocument_HybridSplitsTablesWithinPerSlideCaps() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(520),
                PageHeight = global::OfficeIMO.OfficeLength.Points(420),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(36),
                MarginRight = global::OfficeIMO.OfficeLength.Points(36),
                MarginTop = global::OfficeIMO.OfficeLength.Points(36),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(36),
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(9)
            })
            .Table(new[] {
                new[] { "C1", "C2", "C3", "C4" },
                new[] { "R1C1", "R1C2", "R1C3", "R1C4" },
                new[] { "R2C1", "R2C2", "R2C3", "R2C4" },
                new[] { "R3C1", "R3C2", "R3C3", "R3C4" },
                new[] { "R4C1", "R4C2", "R4C3", "R4C4" }
            }, style: new PdfCore.PdfTableStyle {
                ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(80), global::OfficeIMO.OfficeLength.Points(80), global::OfficeIMO.OfficeLength.Points(80), global::OfficeIMO.OfficeLength.Points(80) },
                HeaderRowCount = 1,
                CellPaddingX = global::OfficeIMO.OfficeLength.Points(4),
                CellPaddingY = global::OfficeIMO.OfficeLength.Points(3)
            })
            .ToBytes();
        var options = PowerPointPdfImportOptions.CreateHybrid();
        options.MaxRowsPerSlide = 2;
        options.MaxColumnsPerSlide = 2;

        PdfPowerPointConversionResult result = PdfCore.PdfDocument.Load(pdf).ConvertToPowerPointWithEvidence(options);

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
    public void PdfDocument_ToPowerPointDocument_HybridChargesRepeatedBackgroundsToAggregateBudget() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(520),
                PageHeight = global::OfficeIMO.OfficeLength.Points(420),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(36),
                MarginRight = global::OfficeIMO.OfficeLength.Points(36),
                MarginTop = global::OfficeIMO.OfficeLength.Points(36),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(36),
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(9)
            })
            .Table(new[] {
                new[] { "C1", "C2" },
                new[] { "R1C1", "R1C2" },
                new[] { "R2C1", "R2C2" },
                new[] { "R3C1", "R3C2" }
            }, style: new PdfCore.PdfTableStyle {
                ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(120), global::OfficeIMO.OfficeLength.Points(120) },
                HeaderRowCount = 1
            })
            .ToBytes();
        PdfCore.PdfDocument source = PdfCore.PdfDocument.Load(pdf);
        PdfCore.PdfPageRenderResult rendered = Assert.Single(source.Render.Pages(
            options: new PdfCore.PdfPageRenderOptions { Dpi = 144 }));
        long oneBackgroundBytes = Assert.IsType<byte[]>(rendered.Bytes).LongLength;
        var options = PowerPointPdfImportOptions.CreateHybrid();
        options.MaxRowsPerSlide = 1;
        options.MaxTotalOutputBytes = oneBackgroundBytes * 2;

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            source.ConvertToPowerPointWithEvidence(options));

        Assert.Contains("aggregate output byte limit", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfDocument_ToPowerPointDocument_HybridUsesVisualPlacementForMixedPageAspectRatios() {
        byte[] source = PdfCore.PdfDocument.Create(new PdfCore.PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(420),
                PageHeight = global::OfficeIMO.OfficeLength.Points(240),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(24),
                MarginRight = global::OfficeIMO.OfficeLength.Points(24),
                MarginTop = global::OfficeIMO.OfficeLength.Points(24),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(24),
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(9)
            })
            .Paragraph(p => p.Text("Landscape reference"))
            .PageBreak()
            .Table(new[] {
                new[] { "Code", "Qty" },
                new[] { "A-100", "2" },
                new[] { "B-200", "14" }
            }, style: new PdfCore.PdfTableStyle {
                HeaderRowCount = 1,
                ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(100), global::OfficeIMO.OfficeLength.Points(60) },
                CellPaddingX = global::OfficeIMO.OfficeLength.Points(4),
                CellPaddingY = global::OfficeIMO.OfficeLength.Points(3)
            })
            .ToBytes();
        PdfCore.PdfDocument resized = PdfCore.PdfDocument.Load(source).Pages.SetMediaBox(0, 0, 240, 420, 2);

        PdfPowerPointConversionResult result = resized.ConvertToPowerPointWithEvidence(PowerPointPdfImportOptions.CreateHybrid());

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
    public void PdfDocument_ToPowerPointDocument_HybridRetainsEditableTablesWhenVisualRenderFails() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(420),
                PageHeight = global::OfficeIMO.OfficeLength.Points(360),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(36),
                MarginRight = global::OfficeIMO.OfficeLength.Points(36),
                MarginTop = global::OfficeIMO.OfficeLength.Points(36),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(36),
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
            })
            .Table(new[] {
                new[] { "Code", "Qty" },
                new[] { "A-100", "2" },
                new[] { "B-200", "14" }
            }, style: new PdfCore.PdfTableStyle {
                HeaderRowCount = 1,
                ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(180), global::OfficeIMO.OfficeLength.Points(80) }
            })
            .ToBytes();
        var options = PowerPointPdfImportOptions.CreateHybrid();
        options.MaxPixelsPerPage = 10;

        PdfPowerPointConversionResult result = PdfCore.PdfDocument.Load(pdf).ConvertToPowerPointWithEvidence(options);

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
            global::OfficeIMO.Drawing.OfficeImageExportFormat.Png,
            new byte[] { 1 },
            1,
            1,
            TimeSpan.Zero,
            Array.Empty<PdfCore.PdfRenderCapabilityDiagnostic>(),
            new PdfCore.PdfPageRenderOptions().MaxDiagnosticCharactersPerPage);
        var failedRender = new PdfCore.PdfPageRenderResult(
            2,
            global::OfficeIMO.Drawing.OfficeImageExportFormat.Png,
            null,
            0,
            0,
            TimeSpan.Zero,
            Array.Empty<PdfCore.PdfRenderCapabilityDiagnostic>(),
            new PdfCore.PdfPageRenderOptions().MaxDiagnosticCharactersPerPage,
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
            formFieldCount: 0,
            hasAcroFormXfa: false,
            annotationCount: 0,
            pageActionCount: 0,
            catalogActionCount: 0,
            hasOpenAction: false,
            documentActionCount: 0,
            optionalContentGroupCount: 1,
            pagesWithOptionalContent: 1,
            interactiveMediaAnnotationCount: 0,
            unplacedFormFieldCount: 0,
            outlineCount: 0,
            attachmentCount: 0,
            hasTaggedContent: false,
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
            formFieldCount: 0,
            hasAcroFormXfa: false,
            annotationCount: 0,
            pageActionCount: 0,
            catalogActionCount: 0,
            hasOpenAction: false,
            documentActionCount: 0,
            optionalContentGroupCount: 0,
            pagesWithOptionalContent: 0,
            interactiveMediaAnnotationCount: 0,
            unplacedFormFieldCount: 0,
            outlineCount: 0,
            attachmentCount: 0,
            hasTaggedContent: false,
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
        PdfCore.PdfConversionWarning groupWarning = Assert.Single(
            report.Warnings,
            static warning => warning.Code == "pdf-groups-not-editable");
        Assert.Equal(PdfCore.PdfConversionWarningSeverity.Information, groupWarning.Severity);
        Assert.Equal("VisualOnly", groupWarning.Details["Disposition"]);
        PdfCore.PdfConversionWarning imageWarning = Assert.Single(
            report.Warnings,
            static warning => warning.Code == "pdf-images-not-editable");
        Assert.Equal(PdfCore.PdfConversionWarningSeverity.Information, imageWarning.Severity);
        Assert.Equal("VisualOnly", imageWarning.Details["Disposition"]);
    }

    [Theory]
    [InlineData(1, 0, "pdf-page-actions-not-reconstructed", OfficeConversionLossKind.Omission, true)]
    [InlineData(0, 1, "pdf-groups-not-editable", OfficeConversionLossKind.Approximation, false)]
    [InlineData(0, 0, "pdf-navigation-not-editable", OfficeConversionLossKind.Approximation, false)]
    [InlineData(0, 0, "pdf-forms-and-controls-not-editable", OfficeConversionLossKind.Approximation, false)]
    [InlineData(0, 0, "pdf-annotations-not-editable", OfficeConversionLossKind.Approximation, false)]
    [InlineData(0, 0, "pdf-animations-not-editable", OfficeConversionLossKind.Approximation, false)]
    public void PdfPowerPointConversionReport_HybridReportsNonvisualPageFactsAsLoss(
        int pageActionCount,
        int optionalContentPageCount,
        string warningCode,
        OfficeConversionLossKind lossKind,
        bool hasOmittedPageContent) {
        var successfulRender = new PdfCore.PdfPageRenderResult(
            1, global::OfficeIMO.Drawing.OfficeImageExportFormat.Png, new byte[] { 1 }, 1, 1, TimeSpan.Zero,
            Array.Empty<PdfCore.PdfRenderCapabilityDiagnostic>(),
            new PdfCore.PdfPageRenderOptions().MaxDiagnosticCharactersPerPage);
        var scope = new PdfCore.PdfTableExtractionScopeReport(
            sourcePageCount: 1, pagesWithTables: 1, detectedTableCount: 1,
            nonTableTextBlockCount: 0, vectorPrimitiveCount: 0, imageCount: 0,
            linkCount: warningCode == "pdf-navigation-not-editable" ? 1 : 0,
            formWidgetCount: warningCode == "pdf-forms-and-controls-not-editable" ? 1 : 0,
            formFieldCount: 0, hasAcroFormXfa: false,
            annotationCount: warningCode == "pdf-annotations-not-editable" ? 1 : 0,
            pageActionCount: pageActionCount,
            catalogActionCount: 0, hasOpenAction: false, documentActionCount: 0,
            optionalContentGroupCount: optionalContentPageCount,
            pagesWithOptionalContent: optionalContentPageCount,
            interactiveMediaAnnotationCount: warningCode == "pdf-animations-not-editable" ? 1 : 0,
            unplacedFormFieldCount: 0,
            outlineCount: 0, attachmentCount: 0, hasTaggedContent: false,
            analysisTruncated: false);
        var report = new PdfPowerPointConversionReport(
            Array.Empty<PdfPowerPointTableImportEntry>(),
            new[] { new PdfPowerPointVisualPageEntry(successfulRender, slideIndex: 0) },
            scope, scope);

        Assert.True(report.HasLoss);
        Assert.Equal(hasOmittedPageContent, report.HasOmittedPageContent);
        PdfCore.PdfConversionWarning warning = Assert.Single(report.Warnings,
            item => item.Code == warningCode);
        Assert.Equal(lossKind, warning.LossKind);
        if (warningCode == "pdf-groups-not-editable") Assert.Equal("VisualOnly", warning.Details["Disposition"]);
    }

    [Fact]
    public void PdfPowerPointConversionReport_ClassifiesOmittedEditableTableContentAsOmission() {
        byte[] pdf = PdfCore.PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Text outside a table"))
            .ToBytes();
        PdfPowerPointConversionResult result = PdfCore.PdfDocument.Load(pdf)
            .ConvertToPowerPointWithEvidence(PowerPointPdfImportOptions.CreateEditableTables());
        using (result.Value) {
            PdfCore.PdfConversionWarning warning = Assert.Single(result.Warnings,
                static item => item.Code == "pdf-text-not-editable");
            Assert.Equal("Omitted", warning.Details["Disposition"]);
            Assert.Equal(OfficeConversionLossKind.Omission, warning.LossKind);
            Assert.True(result.HasOmittedPageContent);
        }
    }

    [Fact]
    public void PdfPowerPointConversionReport_ClassifiesPartiallyOmittedHybridContentAsOmission() {
        var successfulRender = new PdfCore.PdfPageRenderResult(
            1, global::OfficeIMO.Drawing.OfficeImageExportFormat.Png, new byte[] { 1 }, 1, 1, TimeSpan.Zero,
            Array.Empty<PdfCore.PdfRenderCapabilityDiagnostic>(),
            new PdfCore.PdfPageRenderOptions().MaxDiagnosticCharactersPerPage);
        var failedRender = new PdfCore.PdfPageRenderResult(
            2, global::OfficeIMO.Drawing.OfficeImageExportFormat.Png, null, 0, 0, TimeSpan.Zero,
            Array.Empty<PdfCore.PdfRenderCapabilityDiagnostic>(),
            new PdfCore.PdfPageRenderOptions().MaxDiagnosticCharactersPerPage,
            new[] { "render failed" });
        var sourceScope = new PdfCore.PdfTableExtractionScopeReport(
            sourcePageCount: 2, pagesWithTables: 0, detectedTableCount: 0,
            nonTableTextBlockCount: 2, vectorPrimitiveCount: 0, imageCount: 0,
            linkCount: 0, formWidgetCount: 0, formFieldCount: 0, hasAcroFormXfa: false,
            annotationCount: 0, pageActionCount: 0, catalogActionCount: 0,
            hasOpenAction: false, documentActionCount: 0, optionalContentGroupCount: 0,
            pagesWithOptionalContent: 0, interactiveMediaAnnotationCount: 0,
            unplacedFormFieldCount: 0, outlineCount: 0, attachmentCount: 0,
            hasTaggedContent: false, analysisTruncated: false);
        var failedScope = new PdfCore.PdfTableExtractionScopeReport(
            sourcePageCount: 1, pagesWithTables: 0, detectedTableCount: 0,
            nonTableTextBlockCount: 1, vectorPrimitiveCount: 0, imageCount: 0,
            linkCount: 0, formWidgetCount: 0, formFieldCount: 0, hasAcroFormXfa: false,
            annotationCount: 0, pageActionCount: 0, catalogActionCount: 0,
            hasOpenAction: false, documentActionCount: 0, optionalContentGroupCount: 0,
            pagesWithOptionalContent: 0, interactiveMediaAnnotationCount: 0,
            unplacedFormFieldCount: 0, outlineCount: 0, attachmentCount: 0,
            hasTaggedContent: false, analysisTruncated: false);
        var report = new PdfPowerPointConversionReport(
            Array.Empty<PdfPowerPointTableImportEntry>(),
            new[] {
                new PdfPowerPointVisualPageEntry(successfulRender, slideIndex: 0),
                new PdfPowerPointVisualPageEntry(failedRender, slideIndex: 1)
            },
            sourceScope,
            failedScope);

        PdfCore.PdfConversionWarning warning = Assert.Single(report.Warnings,
            static item => item.Code == "pdf-text-not-editable");
        Assert.Equal("PartiallyOmitted", warning.Details["Disposition"]);
        Assert.Equal(OfficeConversionLossKind.Omission, warning.LossKind);
        Assert.True(report.HasOmittedPageContent);
    }

    [Fact]
    public void PdfTables_SaveTablesAsPowerPoint_ImportsDetectedTablesAsPowerPointTables() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(420),
                PageHeight = global::OfficeIMO.OfficeLength.Points(360),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(36),
                MarginRight = global::OfficeIMO.OfficeLength.Points(36),
                MarginTop = global::OfficeIMO.OfficeLength.Points(36),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(36),
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
            })
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "A-100", "Alpha", "2" },
                new[] { "B-200", "Beta", "14" }
            }, style: new PdfCore.PdfTableStyle {
                ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(70), global::OfficeIMO.OfficeLength.Points(170), global::OfficeIMO.OfficeLength.Points(60) },
                HeaderRowCount = 1,
                CellPaddingX = global::OfficeIMO.OfficeLength.Points(6),
                CellPaddingY = global::OfficeIMO.OfficeLength.Points(4)
            })
            .ToBytes();

        using var presentation = new MemoryStream();
        PdfPowerPointConversionResult converted = LoadTables(pdf).ConvertToPowerPointWithEvidence(PowerPointPdfImportOptions.CreateEditableTables());
        using (converted.Value) converted.Value.Save(presentation);
        PdfPowerPointConversionReport report = converted.Report;

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
    public void PdfDocument_ToPowerPointDocument_VisualReportsFailedPagePixelsAsOmitted() {
        byte[] pdf = PdfCore.PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Invoice"))
            .ToBytes();
        var options = PowerPointPdfImportOptions.CreateVisualPages();
        options.MaxPixelsPerPage = 10;

        PdfPowerPointConversionResult result = PdfCore.PdfDocument.Load(pdf).ConvertToPowerPointWithEvidence(options);
        using (result.Value) {
            Assert.False(Assert.Single(result.Report.VisualPages).Succeeded);
            Assert.True(result.Report.HasOmittedPageContent);
            Assert.Contains(result.Report.Warnings, static warning =>
                warning.Code == "pdf-visual-page-render-failed" &&
                warning.LossKind == OfficeConversionLossKind.Omission &&
                warning.Details["Disposition"] == "Omitted");
        }
    }

    [Fact]
    public void PdfDocument_ToPowerPointDocument_VisualScopesOutlineWarningsToSelectedPages() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfCreateOptions { CreateOutlineFromHeadings = true })
            .Paragraph(paragraph => paragraph.Text("Selected first page"))
            .PageBreak()
            .H1("Excluded second-page outline")
            .ToBytes();
        PdfCore.PdfDocument opened = PdfCore.PdfDocument.Load(pdf);
        PdfCore.PdfDocumentInfo sourceInfo = opened.Inspect();
        Assert.Single(sourceInfo.Outlines);
        Assert.Equal(2, sourceInfo.Outlines[0].PageNumber);
        var options = PowerPointPdfImportOptions.CreateVisualPages();
        options.ReadOptions = new PdfCore.PdfLoadOptions {
            PageSelection = PdfCore.PdfPageSelection.From(1)
        };

        PdfPowerPointConversionResult result = opened.ConvertToPowerPointWithEvidence(options);

        using (result.Value) {
            Assert.Equal(1, Assert.Single(result.Report.VisualPages).PageNumber);
            Assert.DoesNotContain(result.Warnings, static warning =>
                warning.Code == "pdf-outlines-not-reconstructed");
        }
    }

    [Fact]
    public void PdfEditableContent_PreservesTypographyAcrossUserUnitScaling() {
        byte[] source = PdfCore.PdfDocument.Create(new PdfCore.PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(420),
                PageHeight = global::OfficeIMO.OfficeLength.Points(360),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(36),
                MarginRight = global::OfficeIMO.OfficeLength.Points(36),
                MarginTop = global::OfficeIMO.OfficeLength.Points(36),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(36),
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
            })
            .Paragraph(paragraph => paragraph.Text("Standalone marker"))
            .Table(new[] {
                new[] { "Code", "Qty" },
                new[] { "A-100", "2" }
            }, style: new PdfCore.PdfTableStyle {
                ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(150), global::OfficeIMO.OfficeLength.Points(80) },
                HeaderRowCount = 1,
                CellPaddingX = global::OfficeIMO.OfficeLength.Points(6),
                CellPaddingY = global::OfficeIMO.OfficeLength.Points(4)
            })
            .ToBytes();

        PdfPowerPointConversionResult baseline = LoadTables(source).ConvertToPowerPointWithEvidence(
            PowerPointPdfImportOptions.CreateEditableContent());
        PdfPowerPointConversionResult scaled = LoadTables(WithUserUnit(source, 2D)).ConvertToPowerPointWithEvidence(
            PowerPointPdfImportOptions.CreateEditableContent());

        using (baseline.Value)
        using (scaled.Value) {
            OfficeIMO.PowerPoint.PowerPointSlide baselineSlide = Assert.Single(baseline.Value.Slides);
            OfficeIMO.PowerPoint.PowerPointSlide scaledSlide = Assert.Single(scaled.Value.Slides);
            double? baselineTableFontSize = Assert.Single(baselineSlide.Tables).GetCell(0, 0).FontSize?.ToPoints();
            double? scaledTableFontSize = Assert.Single(scaledSlide.Tables).GetCell(0, 0).FontSize?.ToPoints();
            double? baselineTextFontSize = Assert.Single(
                baselineSlide.TextBoxes,
                static textBox => textBox.Text == "Standalone marker").Paragraphs[0].Runs[0].FontSize?.ToPoints();
            double? scaledTextFontSize = Assert.Single(
                scaledSlide.TextBoxes,
                static textBox => textBox.Text == "Standalone marker").Paragraphs[0].Runs[0].FontSize?.ToPoints();

            Assert.Equal(baselineTableFontSize, scaledTableFontSize);
            Assert.Equal(baselineTextFontSize, scaledTextFontSize);
        }
    }

    [Fact]
    public void PdfEditableContent_PreservesTableTypographyAboveEighteenPoints() {
        byte[] source = PdfCore.PdfDocument.Create(new PdfCore.PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(420),
                PageHeight = global::OfficeIMO.OfficeLength.Points(360),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(36),
                MarginRight = global::OfficeIMO.OfficeLength.Points(36),
                MarginTop = global::OfficeIMO.OfficeLength.Points(36),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(36),
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(24)
            })
            .Table(new[] {
                new[] { "Metric", "Value" },
                new[] { "Large", "24 pt" }
            }, style: new PdfCore.PdfTableStyle {
                FontSize = global::OfficeIMO.OfficeLength.Points(24),
                HeaderFontSize = global::OfficeIMO.OfficeLength.Points(24),
                ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(150), global::OfficeIMO.OfficeLength.Points(100) },
                HeaderRowCount = 1
            })
            .ToBytes();

        PdfPowerPointConversionResult result = LoadTables(source).ConvertToPowerPointWithEvidence(
            PowerPointPdfImportOptions.CreateEditableContent());
        byte[] serialized;
        int tableFontSize;
        using (result.Value) {
            OfficeIMO.PowerPoint.PowerPointSlide slide = Assert.Single(result.Value.Slides);
            tableFontSize = (int)(Assert.Single(slide.Tables).GetCell(0, 0).FontSize?.ToPoints() ?? 0);
            Assert.True(tableFontSize > 18);
            serialized = result.Value.ToBytes();
        }

        using OfficeIMO.PowerPoint.PowerPointDocument reopened =
            OfficeIMO.PowerPoint.PowerPointDocument.Load(new MemoryStream(serialized));
        OfficeIMO.PowerPoint.PowerPointSlide reopenedSlide = Assert.Single(reopened.Slides);
        int reopenedTableFontSize = (int)(Assert.Single(reopenedSlide.Tables).GetCell(0, 0).FontSize?.ToPoints() ?? 0);
        Assert.True(reopenedTableFontSize > 18);
        Assert.Equal(tableFontSize, reopenedTableFontSize);
    }

    [Fact]
    public void PdfTables_SaveTablesAsPowerPoint_AppliesRowCapsAndKeepsPresentationValidWhenEmpty() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(420),
                PageHeight = global::OfficeIMO.OfficeLength.Points(360),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(36),
                MarginRight = global::OfficeIMO.OfficeLength.Points(36),
                MarginTop = global::OfficeIMO.OfficeLength.Points(36),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(36),
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
            })
            .KeyValueTable(new[] {
                PdfCore.PdfKeyValueRow.Text("InvoiceId", "INV-001"),
                PdfCore.PdfKeyValueRow.Text("Customer", "Evotec"),
                PdfCore.PdfKeyValueRow.Text("Due", "2026-06-30")
            }, style: new PdfCore.PdfTableStyle {
                ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(120), global::OfficeIMO.OfficeLength.Points(170) },
                CellPaddingX = global::OfficeIMO.OfficeLength.Points(6),
                CellPaddingY = global::OfficeIMO.OfficeLength.Points(4)
            })
            .PageBreak()
            .Paragraph(p => p.Text("No table on this page."))
            .ToBytes();

        using var presentation = new MemoryStream();
        PdfPowerPointConversionResult converted = LoadTables(pdf, PdfCore.PdfPageRange.From(1, 1)).ConvertToPowerPointWithEvidence(new PowerPointPdfImportOptions {
                Mode = PowerPointPdfImportMode.EditableTables,
                MaxRows = 2,
                IncludeSourceTitles = false
            });
        using (converted.Value) converted.Value.Save(presentation);
        PdfPowerPointConversionReport report = converted.Report;

        PdfPowerPointTableImportEntry result = Assert.Single(report.TableEntries);
        Assert.Equal(2, result.RowCount);
        Assert.Equal(3, result.TotalRowCount);
        Assert.True(result.Truncated);
        Assert.False(result.HeaderRowIncluded);
        Assert.True(report.HasLoss);
        Assert.Throws<InvalidOperationException>(() => report.RequireNoLoss());

        using (PresentationDocument package = PresentationDocument.Open(new MemoryStream(presentation.ToArray()), false)) {
            Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
            A.Table table = GetSingleTable(package);
            Assert.False(table.TableProperties?.FirstRow?.Value ?? false);
            List<A.TableRow> rows = table.Elements<A.TableRow>().ToList();
            Assert.Equal(2, rows.Count);
            Assert.Equal(new[] { "InvoiceId", "INV-001" }, ReadRowText(rows[0]));
            Assert.Equal(new[] { "Customer", "Evotec" }, ReadRowText(rows[1]));
        }

        using var emptyPresentation = new MemoryStream();
        PdfPowerPointConversionResult convertedEmpty = LoadTables(pdf, PdfCore.PdfPageRange.From(2, 2)).ConvertToPowerPointWithEvidence(new PowerPointPdfImportOptions {
                Mode = PowerPointPdfImportMode.EditableTables,
                EmptyPresentationMessage = "Nothing tabular was detected."
            });
        using (convertedEmpty.Value) convertedEmpty.Value.Save(emptyPresentation);
        PdfPowerPointConversionReport emptyReport = convertedEmpty.Report;

        Assert.Empty(emptyReport.TableEntries);
        using PresentationDocument emptyPackage = PresentationDocument.Open(new MemoryStream(emptyPresentation.ToArray()), false);
        Assert.Empty(new OpenXmlValidator().Validate(emptyPackage).ToList());
        Assert.Empty(emptyPackage.PresentationPart!.SlideParts.SelectMany(part => part.Slide.Descendants<A.Table>()));
        Assert.Contains(ReadAllText(emptyPackage), text => text == "Nothing tabular was detected.");
    }

    [Fact]
    public void PdfTables_SaveTablesAsPowerPoint_SkipsHeaderOnlySegmentsWhenHeadersAreDisabled() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(420),
                PageHeight = global::OfficeIMO.OfficeLength.Points(260),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(36),
                MarginRight = global::OfficeIMO.OfficeLength.Points(36),
                MarginTop = global::OfficeIMO.OfficeLength.Points(36),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(36),
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
            })
            .Table(new[] {
                new[] { "Code", "Name" }
            }, style: new PdfCore.PdfTableStyle {
                ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(90), global::OfficeIMO.OfficeLength.Points(150) },
                HeaderRowCount = 1,
                CellPaddingX = global::OfficeIMO.OfficeLength.Points(6),
                CellPaddingY = global::OfficeIMO.OfficeLength.Points(4)
            })
            .ToBytes();

        using var presentation = new MemoryStream();
        PdfPowerPointConversionResult converted = LoadTables(pdf).ConvertToPowerPointWithEvidence(new PowerPointPdfImportOptions {
                Mode = PowerPointPdfImportMode.EditableTables,
                IncludeColumnHeaderRows = false,
                EmptyPresentationMessage = "No table rows were imported."
            });
        using (converted.Value) converted.Value.Save(presentation);
        PdfPowerPointConversionReport report = converted.Report;

        Assert.Empty(report.TableEntries);
        using PresentationDocument package = PresentationDocument.Open(new MemoryStream(presentation.ToArray()), false);
        Assert.Empty(new OpenXmlValidator().Validate(package).ToList());
        Assert.Empty(package.PresentationPart!.SlideParts.SelectMany(part => part.Slide.Descendants<A.Table>()));
        Assert.Contains(ReadAllText(package), text => text == "No table rows were imported.");
    }

    [Fact]
    public void PdfTables_SaveTablesAsPowerPoint_SplitsLargeTablesAcrossSlides() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(520),
                PageHeight = global::OfficeIMO.OfficeLength.Points(420),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(36),
                MarginRight = global::OfficeIMO.OfficeLength.Points(36),
                MarginTop = global::OfficeIMO.OfficeLength.Points(36),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(36),
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(9)
            })
            .Table(new[] {
                new[] { "C1", "C2", "C3", "C4" },
                new[] { "R1C1", "R1C2", "R1C3", "R1C4" },
                new[] { "R2C1", "R2C2", "R2C3", "R2C4" },
                new[] { "R3C1", "R3C2", "R3C3", "R3C4" },
                new[] { "R4C1", "R4C2", "R4C3", "R4C4" }
            }, style: new PdfCore.PdfTableStyle {
                ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(80), global::OfficeIMO.OfficeLength.Points(80), global::OfficeIMO.OfficeLength.Points(80), global::OfficeIMO.OfficeLength.Points(80) },
                HeaderRowCount = 1,
                CellPaddingX = global::OfficeIMO.OfficeLength.Points(4),
                CellPaddingY = global::OfficeIMO.OfficeLength.Points(3)
            })
            .ToBytes();

        using var presentation = new MemoryStream();
        PdfPowerPointConversionResult converted = LoadTables(pdf).ConvertToPowerPointWithEvidence(new PowerPointPdfImportOptions {
                Mode = PowerPointPdfImportMode.EditableTables,
                MaxRowsPerSlide = 2,
                MaxColumnsPerSlide = 2
            });
        using (converted.Value) converted.Value.Save(presentation);
        PdfPowerPointConversionReport report = converted.Report;

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
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(320),
                PageHeight = global::OfficeIMO.OfficeLength.Points(220),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
                MarginRight = global::OfficeIMO.OfficeLength.Points(30),
                MarginTop = global::OfficeIMO.OfficeLength.Points(30),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(30),
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(9)
            })
            .Table(rows, style: new PdfCore.PdfTableStyle {
                HeaderRowCount = 2,
                RepeatHeaderRowCount = 2,
                ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(120), global::OfficeIMO.OfficeLength.Points(120) },
                CellPaddingX = global::OfficeIMO.OfficeLength.Points(5),
                CellPaddingY = global::OfficeIMO.OfficeLength.Points(3)
            })
            .ToBytes();

        PdfPowerPointConversionResult result = LoadTables(pdf).ConvertToPowerPointWithEvidence(new PowerPointPdfImportOptions {
            Mode = PowerPointPdfImportMode.EditableTables,
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

    private static byte[] WithUserUnit(byte[] source, double userUnit) =>
        PdfCore.PdfDocumentObjectGraphRewriter.Rewrite(source, null, null, (objects, security) => {
            PdfCore.PdfIndirectObject page = Assert.Single(objects.Values, static item =>
                item.Value is PdfCore.PdfDictionary dictionary &&
                string.Equals(dictionary.Get<PdfCore.PdfName>("Type")?.Name, "Page", StringComparison.Ordinal));
            Assert.IsType<PdfCore.PdfDictionary>(page.Value).Items["UserUnit"] = new PdfCore.PdfNumber(userUnit);
            return security.InfoObjectNumber;
        });

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
