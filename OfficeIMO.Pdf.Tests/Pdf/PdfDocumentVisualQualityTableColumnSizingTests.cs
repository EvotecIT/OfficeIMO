using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Fact]
    public void Table_UsesFixedColumnWidthPointsWithRemainingWeightedColumns() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.ColumnWidthPoints = new List<double?> { 60, null, 50 };

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] { "ID", "Description", "Score" },
                new[] { "A1", "Longer descriptive value", "100" }
            }, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double idX = FindWordStartX(page, "ID");
        double descriptionX = FindWordStartX(page, "Description");
        double scoreX = FindWordStartX(page, "Score");

        double firstColumnWidth = descriptionX - idX;
        double secondColumnWidth = scoreX - descriptionX;
        Assert.InRange(firstColumnWidth, 55, 65);
        Assert.True(secondColumnWidth > 170, $"Expected the unfixed middle table column to consume remaining width. Second gap: {secondColumnWidth:0.##}.");
    }

    [Fact]
    public void RowColumnTable_UsesFixedColumnWidthPointsWithRemainingWeightedColumns() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.ColumnWidthPoints = new List<double?> { 60, null, 50 };

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "ID", "Description", "Score" },
                                    new[] { "A1", "Longer descriptive value", "100" }
                                }, style: style))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double idX = FindWordStartX(page, "ID");
        double descriptionX = FindWordStartX(page, "Description");
        double scoreX = FindWordStartX(page, "Score");

        double firstColumnWidth = descriptionX - idX;
        double secondColumnWidth = scoreX - descriptionX;
        Assert.InRange(firstColumnWidth, 55, 65);
        Assert.True(secondColumnWidth > 170, $"Expected the row-column unfixed middle table column to consume remaining width. Second gap: {secondColumnWidth:0.##}.");
    }

    [Fact]
    public void Table_UsesMinimumColumnWidthPointsForWeightedColumns() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.ColumnWidthWeights = new List<double> { 1, 10, 1 };
        style.ColumnMinWidthPoints = new List<double?> { 80, null, null };

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] { "ID", "Description", "Score" },
                new[] { "A1", "Longer descriptive value", "100" }
            }, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double idX = FindWordStartX(page, "ID");
        double descriptionX = FindWordStartX(page, "Description");
        double firstColumnWidth = descriptionX - idX;

        Assert.InRange(firstColumnWidth, 75, 85);
    }

    [Fact]
    public void RowColumnTable_UsesMinimumColumnWidthPointsForWeightedColumns() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.ColumnWidthWeights = new List<double> { 1, 10, 1 };
        style.ColumnMinWidthPoints = new List<double?> { 80, null, null };

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "ID", "Description", "Score" },
                                    new[] { "A1", "Longer descriptive value", "100" }
                                }, style: style))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double idX = FindWordStartX(page, "ID");
        double descriptionX = FindWordStartX(page, "Description");
        double firstColumnWidth = descriptionX - idX;

        Assert.InRange(firstColumnWidth, 75, 85);
    }

    [Fact]
    public void Table_ScalesMinimumColumnWidthsWhenTheyExceedAvailableWidth() {
        var options = new PdfOptions {
            PageWidth = 300,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.ColumnWidthWeights = new List<double> { 1, 1, 1 };
        style.ColumnMinWidthPoints = new List<double?> { 180, 180, 180 };

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] { "A", "B", "C" },
                new[] { "1", "2", "3" }
            }, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("A", text);
        Assert.Contains("B", text);
        Assert.Contains("C", text);
    }

    [Fact]
    public void Table_UsesMaximumColumnWidthPointsForWeightedColumns() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.ColumnWidthWeights = new List<double> { 1, 10, 1 };
        style.ColumnMaxWidthPoints = new List<double?> { null, 120, null };

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] { "ID", "Description", "Score" },
                new[] { "A1", "Longer descriptive value", "100" }
            }, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double descriptionX = FindWordStartX(page, "Description");
        double scoreX = FindWordStartX(page, "Score");
        double secondColumnWidth = scoreX - descriptionX;

        Assert.InRange(secondColumnWidth, 115, 125);
    }

    [Fact]
    public void RowColumnTable_UsesMaximumColumnWidthPointsForWeightedColumns() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.ColumnWidthWeights = new List<double> { 1, 10, 1 };
        style.ColumnMaxWidthPoints = new List<double?> { null, 120, null };

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "ID", "Description", "Score" },
                                    new[] { "A1", "Longer descriptive value", "100" }
                                }, style: style))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double descriptionX = FindWordStartX(page, "Description");
        double scoreX = FindWordStartX(page, "Score");
        double secondColumnWidth = scoreX - descriptionX;

        Assert.InRange(secondColumnWidth, 115, 125);
    }

    [Fact]
    public void Table_UsesConfiguredVerticalColumnAlignment() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.ColumnWidthPoints = new List<double?> { 80, null };
        style.VerticalAlignments = new List<PdfCellVerticalAlign> { PdfCellVerticalAlign.Bottom, PdfCellVerticalAlign.Top };

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] { "Name", "Notes" },
                new[] {
                    "BottomValue",
                    "This note wraps across several lines so the row becomes tall enough to make vertical alignment visible in the first cell."
                }
            }, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double bottomValueY = FindWordStartY(page, "BottomValue");
        double wrappedFirstLineY = FindWordStartY(page, "This");

        Assert.True(bottomValueY < wrappedFirstLineY - 10, $"Expected the first-column value to sit lower than the top-aligned wrapped text. BottomValue y: {bottomValueY:0.##}, wrapped y: {wrappedFirstLineY:0.##}.");
    }

    [Fact]
    public void RowColumnTable_UsesConfiguredVerticalColumnAlignment() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.ColumnWidthPoints = new List<double?> { 80, null };
        style.VerticalAlignments = new List<PdfCellVerticalAlign> { PdfCellVerticalAlign.Bottom, PdfCellVerticalAlign.Top };

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "Name", "Notes" },
                                    new[] {
                                        "BottomValue",
                                        "This note wraps across several lines so the row becomes tall enough to make vertical alignment visible in the first cell."
                                    }
                                }, style: style))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double bottomValueY = FindWordStartY(page, "BottomValue");
        double wrappedFirstLineY = FindWordStartY(page, "This");

        Assert.True(bottomValueY < wrappedFirstLineY - 10, $"Expected the first row-column cell value to sit lower than the top-aligned wrapped text. BottomValue y: {bottomValueY:0.##}, wrapped y: {wrappedFirstLineY:0.##}.");
    }

    [Fact]
    public void Table_AutoFitsFlexibleColumnsFromMeasuredContent() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.AutoFitColumns = true;

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] { "SKU", "Description", "Amount" },
                new[] { "A1", "Managed service renewal with monitoring and incident response", "1250" }
            }, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double skuX = FindWordStartX(page, "SKU");
        double descriptionX = FindWordStartX(page, "Description");
        double amountX = FindWordStartX(page, "Amount");
        double firstColumnWidth = descriptionX - skuX;
        double secondColumnWidth = amountX - descriptionX;
        double rightMost = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .Max(letter => letter.EndBaseLine.X);

        Assert.True(secondColumnWidth > firstColumnWidth * 3, $"Expected measured content to make the description column much wider. First gap: {firstColumnWidth:0.##}, second gap: {secondColumnWidth:0.##}.");
        Assert.True(secondColumnWidth > 190, $"Expected measured content to reserve substantial width for the description column. Second gap: {secondColumnWidth:0.##}.");
        Assert.InRange(rightMost, double.NegativeInfinity, options.PageWidth - options.MarginRight + 3);
    }

    [Fact]
    public void Table_AutoFitUsesPreferredWidthForCompactContent() {
        var style = TableStyles.Minimal();
        style.AutoFitColumns = true;
        style.PreferredWidth = 160;
        style.PreserveWidth = true;
        style.HeaderFill = new PdfColor(0.17, 0.27, 0.37);
        style.BorderColor = null;

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 360,
                PageHeight = 240,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFontSize = 9
            })
            .Table(new[] {
                new[] { "ID", "State" }
            }, style: style)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        double renderedWidth = ExtractPaintedRectangles(content, "0.17 0.27 0.37 rg", "f").Sum(rectangle => rectangle.W);

        Assert.InRange(renderedWidth, 159.5D, 160.5D);
    }

    [Fact]
    public void RowColumnTable_AutoFitCanExpandPastPreferredWidthWithinMaximumWidth() {
        var style = TableStyles.Minimal();
        style.AutoFitColumns = true;
        style.PreferredWidth = 80;
        style.MaxWidth = 150;
        style.PreserveWidth = true;
        style.HeaderFill = new PdfColor(0.18, 0.28, 0.38);
        style.BorderColor = null;

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 360,
                PageHeight = 240,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFontSize = 9
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "MMMMMMMMMMMMMMMMMMMMMMMM", "State" }
                                }, style: style))))))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        double renderedWidth = ExtractPaintedRectangles(content, "0.18 0.28 0.38 rg", "f").Sum(rectangle => rectangle.W);

        Assert.InRange(renderedWidth, 100D, 150.5D);
    }

    [Fact]
    public void Table_AutoFitsTechnicalValuesUsingBreakAwarePreferredWidths() {
        var options = new PdfOptions {
            PageWidth = 612,
            PageHeight = 360,
            MarginLeft = 72,
            MarginRight = 72,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.AutoFitColumns = true;
        style.MaxWidth = options.PageWidth - options.MarginLeft - options.MarginRight;

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] { "Item", "IssuedOn", "Category", "Reference", "ClausePath", "UpdatedOn", "Status" },
                new[] {
                    "Invoice | Review 10",
                    "05/26/2023 09:00:56",
                    "commercial.contract",
                    "253cfd36-2f82-4672-b8e3-31b7a8ebaaf4",
                    "Section=Revenue,Article=LateFee,Clause={253CFD36-2F82-4672-B8E3-31B7A8EBAAF4},Page=12,Paragraph=4,Region=Global",
                    "05/26/2023 09:00:56",
                    "Allow"
                }
            }, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double categoryX = FindWordStartX(page, "Category");
        double clausePathX = FindWordStartX(page, "ClausePath");
        double referenceX = FindWordStartX(page, "Reference");
        double statusX = FindWordStartX(page, "Status");
        double accessX = FindWordStartX(page, "Allow");

        Assert.True(referenceX - categoryX > 35D, $"Expected dotted qualified values to receive flexible width instead of collapsing to only their shortest dot-delimited segment. Category width: {referenceX - categoryX:0.##}.");
        Assert.InRange(clausePathX, 250D, 290D);
        Assert.True(clausePathX - referenceX < 70D, $"Expected compact GUID-like references to wrap at separators instead of starving structured path columns. Reference width: {clausePathX - referenceX:0.##}.");
        Assert.True(statusX - clausePathX > 220D, $"Expected structured path columns to receive the broad share of break-aware technical tables. ClausePath width: {statusX - clausePathX:0.##}.");
        Assert.True(accessX > 480D, $"Expected break-aware autofit to reserve a Word-like structured path column from value shape, not header names. Status value left: {accessX:0.##}.");
    }

    [Fact]
    public void Table_AutoFitsRegistryPathColumnsWithoutStarvingPathWidth() {
        var options = new PdfOptions {
            PageWidth = 612,
            PageHeight = 360,
            MarginLeft = 72,
            MarginRight = 72,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.AutoFitColumns = true;
        style.MaxWidth = options.PageWidth - options.MarginLeft - options.MarginRight;

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] { "Setting", "Key", "ValueName", "Effective", "EffectiveSource", "Sources" },
                new[] {
                    "DirectoryClientIntegrity",
                    @"HKLM\SYSTEM\CurrentControlSet\Services\Directory\Security\Parameters",
                    "DirectoryClientIntegrity",
                    "Enabled",
                    "Baseline Settings",
                    "2 items"
                },
                new[] {
                    "DirectoryChannelBinding",
                    @"HKLM\SYSTEM\CurrentControlSet\Services\Directory\Security\ChannelBinding",
                    "DirectoryChannelBinding",
                    "Not configured",
                    "",
                    "0 items"
                }
            }, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double keyX = FindWordStartX(page, "Key");
        double valueNameX = FindWordStartX(page, "ValueName");
        double effectiveX = FindWordStartX(page, "Effective");
        double sourcesX = FindWordStartX(page, "Sources");
        double rightMost = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .Max(letter => letter.EndBaseLine.X);

        Assert.True(valueNameX - keyX > 135D, $"Expected registry/path columns to receive enough width to avoid excessive row height. Key width: {valueNameX - keyX:0.##}.");
        Assert.True(sourcesX - effectiveX > 90D, $"Expected later columns to remain usable after widening the path column. Effective-to-sources gap: {sourcesX - effectiveX:0.##}.");
        Assert.InRange(rightMost, double.NegativeInfinity, options.PageWidth - options.MarginRight + 3D);
    }

    [Fact]
    public void Table_AutoFitsUppercaseDelimitedCodesUsingRealBreakpoints() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.AutoFitColumns = true;

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] { "Code", "Description", "Status" },
                new[] { "CASE-REVIEW-LINE-2026-0001", "Contract clause review", "Open" }
            }, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double codeX = FindWordStartX(page, "Code");
        double descriptionX = FindWordStartX(page, "Description");
        double statusX = FindWordStartX(page, "Status");
        double codeColumnWidth = descriptionX - codeX;

        Assert.InRange(codeColumnWidth, 35D, 95D);
        Assert.True(statusX > descriptionX + 100D, $"Expected uppercase delimited code values to wrap at separators instead of consuming the table. Description left: {descriptionX:0.##}, Status left: {statusX:0.##}.");
    }

    [Fact]
    public void Table_AutoFitsSidLikeIdentifiersAsCompactTechnicalValues() {
        var options = new PdfOptions {
            PageWidth = 612,
            PageHeight = 360,
            MarginLeft = 72,
            MarginRight = 72,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.AutoFitColumns = true;
        style.MaxWidth = options.PageWidth - options.MarginLeft - options.MarginRight;

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] { "Name", "Sid", "Kind", "Right", "Scope", "Review", "Owner" },
                new[] {
                    "Creator Owners",
                    "S-1-5-21-853615985-2870445339-3163598659-520",
                    "WellKnownGroup",
                    "ModifyOwner",
                    "Baseline expected delegation",
                    "Long review note that should keep a useful width instead of being starved by the SID column",
                    "Security Team"
                }
            }, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double sidX = FindWordStartX(page, "Sid");
        double kindX = FindWordStartX(page, "Kind");
        double reviewX = FindWordStartX(page, "Review");
        double ownerX = FindWordStartX(page, "Owner");
        double sidColumnWidth = kindX - sidX;
        double reviewColumnWidth = ownerX - reviewX;
        var words = page.GetWords().Select(word => word.Text).ToArray();
        double rightMost = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .Max(letter => letter.EndBaseLine.X);

        Assert.InRange(sidColumnWidth, 40D, 90D);
        Assert.Contains(words, word => word.StartsWith("S-1-5-21-", StringComparison.Ordinal));
        Assert.DoesNotContain(words, word => word.Contains("853615985-2870445339", StringComparison.Ordinal));
        Assert.True(reviewColumnWidth > 120D, $"Expected SID-like identifiers to stay compact so descriptive columns remain usable. Review width: {reviewColumnWidth:0.##}.");
        Assert.InRange(rightMost, double.NegativeInfinity, options.PageWidth - options.MarginRight + 3D);
    }

    [Fact]
    public void Table_AutoFitsCompactNumericDelimitedIdentifiersUsingDelimiterBreaks() {
        var options = new PdfOptions {
            PageWidth = 612,
            PageHeight = 360,
            MarginLeft = 72,
            MarginRight = 72,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.AutoFitColumns = true;
        style.MaxWidth = options.PageWidth - options.MarginLeft - options.MarginRight;

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] { "DisplayName", "Id", "Type", "Permission", "Flag", "Risk", "Notes" },
                new[] {
                    "Everyone",
                    "S-1-1-0",
                    "WellKnownGroup",
                    "Unknown",
                    "False",
                    "High",
                    "Compact numeric identifiers should break at delimiters instead of reserving a wide technical column"
                },
                new[] {
                    "Self",
                    "S-1-5-10",
                    "WellKnownGroup",
                    "Write",
                    "True",
                    "Medium",
                    "Dense report tables need this column to stay narrow enough for the surrounding descriptive cells"
                }
            }, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double identifierX = FindWordStartX(page, "Id");
        double typeX = FindWordStartX(page, "Type");
        double notesX = FindWordStartX(page, "Notes");
        double identifierColumnWidth = typeX - identifierX;
        var words = page.GetWords().Select(word => word.Text).ToArray();

        Assert.InRange(identifierColumnWidth, 14D, 55D);
        Assert.DoesNotContain(words, word => string.Equals(word, "S-1-1-0", StringComparison.Ordinal));
        Assert.Contains(words, word => string.Equals(word, "S-", StringComparison.Ordinal));
        Assert.True(notesX > typeX + 140D, $"Expected compact numeric-delimited identifiers to leave room for descriptive report columns. Type left: {typeX:0.##}, Notes left: {notesX:0.##}.");
    }

    [Fact]
    public void Table_AutoFitsDelimitedNumericIdentifierListsAsCompactTechnicalValues() {
        var options = new PdfOptions {
            PageWidth = 612,
            PageHeight = 360,
            MarginLeft = 72,
            MarginRight = 72,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.AutoFitColumns = true;
        style.MaxWidth = options.PageWidth - options.MarginLeft - options.MarginRight;

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] { "ScopeName", "Spn", "ServiceClass", "AccountCount", "Accounts", "AccountSids" },
                new[] {
                    "contoso.example",
                    "service/changepw",
                    "service",
                    "2",
                    "svc-account",
                    "S-1-5-21-3661168273-3802070955-2987026695-502; S-1-5-21-853615985-2870445339-3163598659-502"
                }
            }, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var words = page.GetWords().Select(word => word.Text).ToArray();

        double rightMost = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .Max(letter => letter.EndBaseLine.X);

        Assert.Contains(words, word => word.StartsWith("S-1-5-21-", StringComparison.Ordinal));
        Assert.Contains(words, word => string.Equals(word, "service/changepw", StringComparison.Ordinal));
        Assert.Contains(words, word => string.Equals(word, "ServiceClass", StringComparison.Ordinal));
        Assert.Contains(words, word => string.Equals(word, "AccountCount", StringComparison.Ordinal));
        Assert.DoesNotContain(words, word => word.Contains("3661168273-3802070955", StringComparison.Ordinal));
        Assert.InRange(rightMost, double.NegativeInfinity, options.PageWidth - options.MarginRight + 3D);
    }

    [Fact]
    public void Table_AutoFitsRepeatedAccountIdentifiersWithComparableColumnWidths() {
        var options = new PdfOptions {
            PageWidth = 612,
            PageHeight = 360,
            MarginLeft = 72,
            MarginRight = 72,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.AutoFitColumns = true;
        style.MaxWidth = options.PageWidth - options.MarginLeft - options.MarginRight;

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] { "DisplayName", "AdminType", "OperationKind", "CanEdit", "RiskLevel", "Identity", "Sid" },
                new[] {
                    @"CONTOSO\Service_2026Inventory",
                    "NotAdministrative",
                    "Write",
                    "True",
                    "Medium",
                    @"CONTOSO\Service_2026Inventory",
                    "S-1-5-21-853615985-2870445339-3163598659-520"
                },
                new[] {
                    @"FABRIKAM\Service_2026Auditor",
                    "NotAdministrative",
                    "Write",
                    "True",
                    "Medium",
                    @"FABRIKAM\Service_2026Auditor",
                    "S-1-5-21-853615985-2870445339-3163598659-521"
                }
            }, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double displayNameX = FindWordStartX(page, "DisplayName");
        double accountTypeX = FindWordStartX(page, "AdminType");
        double identityX = FindWordStartX(page, "Identity");
        double sidX = FindWordStartX(page, "Sid");
        double displayNameColumnWidth = accountTypeX - displayNameX;
        double identityColumnWidth = sidX - identityX;
        double rightMost = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .Max(letter => letter.EndBaseLine.X);

        Assert.True(displayNameColumnWidth > 95D, $"Expected the leading identifier column to remain usable in dense autofit tables. DisplayName width: {displayNameColumnWidth:0.##}.");
        Assert.True(identityColumnWidth > 95D, $"Expected the repeated identifier column to remain usable in dense autofit tables. Identity width: {identityColumnWidth:0.##}.");
        Assert.True(
            identityColumnWidth >= displayNameColumnWidth * 0.65D,
            $"Expected repeated identifier columns to receive comparable width. DisplayName width: {displayNameColumnWidth:0.##}, Identity width: {identityColumnWidth:0.##}.");
        Assert.InRange(rightMost, double.NegativeInfinity, options.PageWidth - options.MarginRight + 3D);
    }

    [Fact]
    public void Table_AutoFitsDenseInventoryColumnsWithoutOverweightingCompactCodes() {
        var options = new PdfOptions {
            PageWidth = 612,
            PageHeight = 360,
            MarginLeft = 72,
            MarginRight = 72,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.AutoFitColumns = true;
        style.MaxWidth = options.PageWidth - options.MarginLeft - options.MarginRight;

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] { "Status", "Priority", "CaseId", "Category", "Region", "EvidencePath", "Owner" },
                new[] {
                    "NonConformingInventoryItem",
                    "High",
                    "CASE-REVIEW-2026-0001",
                    "Policy Review",
                    "north-region",
                    "/contracts/2026/master-services/addendum/section-12/liability",
                    "LegalTeamA"
                },
                new[] {
                    "NonConformingInventoryItem",
                    "High",
                    "CASE-REVIEW-2026-0002",
                    "Policy Review",
                    "south-region",
                    "/operations/locations/site-a/controls/fire-safety/exceptions",
                    "AuditTeamB"
                }
            }, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double evidencePathX = FindWordStartX(page, "EvidencePath");
        double ownerX = FindWordStartX(page, "Owner");
        double rightMost = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .Max(letter => letter.EndBaseLine.X);

        Assert.True(evidencePathX < 335D, $"Expected dense auto-fit to avoid overweighting compact code/header columns before EvidencePath. EvidencePath left: {evidencePathX:0.##}.");
        Assert.True(ownerX - evidencePathX > 150D, $"Expected long path values to receive a Word-like share of dense table width. EvidencePath left: {evidencePathX:0.##}, Owner left: {ownerX:0.##}.");
        Assert.InRange(rightMost, double.NegativeInfinity, options.PageWidth - options.MarginRight + 3D);
    }

    [Fact]
    public void Table_AutoFitsLargeDenseCamelCaseReportsWithoutStarvingCompactColumns() {
        var options = new PdfOptions {
            PageWidth = 612,
            PageHeight = 360,
            MarginLeft = 72,
            MarginRight = 72,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.AutoFitColumns = true;
        style.MaxWidth = options.PageWidth - options.MarginLeft - options.MarginRight;

        var rows = new List<string[]> {
            new[] { "State", "Level", "Code", "Topic", "Zone", "Owner", "Target" }
        };

        for (int index = 1; index <= 110; index++) {
            rows.Add(new[] {
                "MissingRequiredReviewState",
                index % 3 == 0 ? "Medium" : "High",
                "CASE-REVIEW-2026-" + index.ToString("0000", CultureInfo.InvariantCulture),
                index % 2 == 0 ? "Configuration Drift" : "Policy Review",
                index % 2 == 0 ? "NorthRegion" : "SouthRegion",
                index % 2 == 0 ? "OperationsReviewTeam" : "LegalReviewTeam",
                "VeryLongServiceEndpointWithoutSeparatorsThatShouldUseTheRemainingColumnWidth"
            });
        }

        byte[] bytes = PdfDocument.Create(options)
            .Table(rows, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double statusX = FindWordStartX(page, "State");
        double severityX = FindWordStartX(page, "Level");
        double checkIdX = FindWordStartX(page, "Code");
        double targetX = FindWordStartX(page, "Target");
        double rightMost = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .Max(letter => letter.EndBaseLine.X);

        Assert.InRange(severityX - statusX, 45D, 95D);
        Assert.InRange(checkIdX - severityX, 25D, 65D);
        Assert.True(targetX - checkIdX > 150D, $"Expected dense CamelCase auto-fit to preserve room for middle and trailing report columns. Code left: {checkIdX:0.##}, Target left: {targetX:0.##}.");
        Assert.True(pdf.NumberOfPages >= 25, $"Expected large dense auto-fit rows to wrap and paginate instead of compressing into too few pages. Pages: {pdf.NumberOfPages}.");
        Assert.InRange(rightMost, double.NegativeInfinity, options.PageWidth - options.MarginRight + 3D);
    }

    [Fact]
    public void Table_AutoFitsLargeDenseUppercaseCodesWithoutOverprotectingLongSegments() {
        var options = new PdfOptions {
            PageWidth = 612,
            PageHeight = 792,
            MarginLeft = 72,
            MarginRight = 72,
            MarginTop = 72,
            MarginBottom = 72,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 11
        };
        var style = TableStyles.TableGrid();
        style.AutoFitColumns = true;
        style.MaxWidth = options.PageWidth - options.MarginLeft - options.MarginRight;
        style.FontSize = 11D;
        style.LineHeight = 1.22D;
        style.HeaderRowCount = 1;
        style.RepeatHeaderRowCount = 1;
        style.CellPaddingTop = 0D;
        style.CellPaddingBottom = 0D;
        style.CellPaddingLeft = 5.4D;
        style.CellPaddingRight = 5.4D;

        var rows = new List<string[]> {
            new[] { "Status", "Severity", "CheckId", "Category", "ZoneName", "OwnerName", "Target" }
        };

        for (int index = 1; index <= 116; index++) {
            rows.Add(new[] {
                "NonDomainControllerTarget",
                index % 5 == 0 ? "Medium" : "High",
                "ADDC-DNS-LOCATOR-001",
                "AD DC Locator",
                index % 3 == 0 ? "_msdcs.ad.evotec.example" : "test.zone1",
                index % 2 == 0
                    ? "_kerberos._tcp.Default-First-Site-Name._sites.dc._msdcs.ad.evotec.example"
                    : "_ldap._tcp.ForestDnsZones.test.zone1",
                "ad" + (index % 3).ToString(CultureInfo.InvariantCulture) + ".ad.evotec.example"
            });
        }

        byte[] bytes = PdfDocument.Create(options)
            .Table(rows, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var words = page.GetWords().Select(word => word.Text).ToArray();

        Assert.DoesNotContain(words, word => string.Equals(word, "LOCATOR", StringComparison.Ordinal));
        Assert.Contains(words, word => word.StartsWith("LOC", StringComparison.Ordinal));
        Assert.True(pdf.NumberOfPages >= 8, $"Expected dense technical rows to preserve Word-like pagination pressure. Pages: {pdf.NumberOfPages}.");
    }

    [Fact]
    public void RowColumnTable_AutoFitsFlexibleColumnsFromMeasuredContent() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.AutoFitColumns = true;

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "SKU", "Description", "Amount" },
                                    new[] { "A1", "Managed service renewal with monitoring and incident response", "1250" }
                                }, style: style))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double skuX = FindWordStartX(page, "SKU");
        double descriptionX = FindWordStartX(page, "Description");
        double amountX = FindWordStartX(page, "Amount");
        double firstColumnWidth = descriptionX - skuX;
        double secondColumnWidth = amountX - descriptionX;
        double rightMost = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .Max(letter => letter.EndBaseLine.X);

        Assert.True(secondColumnWidth > firstColumnWidth * 3, $"Expected measured content to make the row-column description column much wider. First gap: {firstColumnWidth:0.##}, second gap: {secondColumnWidth:0.##}.");
        Assert.True(secondColumnWidth > 190, $"Expected measured content to reserve substantial width for the row-column description column. Second gap: {secondColumnWidth:0.##}.");
        Assert.InRange(rightMost, double.NegativeInfinity, options.PageWidth - options.MarginRight + 3);
    }

    [Fact]
    public void Table_RightAlignsCurrencyPercentAndParenthesizedNumbers() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.RightAlignedNumbers();
        style.ColumnWidthPoints = new List<double?> { 120, 100 };

        byte[] bytes = PdfDocument.Create(options)
            .Table(new[] {
                new[] { "Metric", "Amount" },
                new[] { "Revenue", "$1,234.50" },
                new[] { "Refund", "(45.20)" },
                new[] { "Margin", "99%" },
                new[] { "EU", "€1,234.50" }
            }, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double dollarEnd = FindWordEndX(page, "$1,234.50");
        double refundEnd = FindWordEndX(page, "(45.20)");
        double percentEnd = FindWordEndX(page, "99%");
        double euroEnd = FindWordEndX(page, "€1,234.50");

        Assert.InRange(Math.Abs(refundEnd - dollarEnd), 0, 3);
        Assert.InRange(Math.Abs(percentEnd - dollarEnd), 0, 3);
        Assert.InRange(Math.Abs(euroEnd - dollarEnd), 0, 3);
    }

    [Fact]
    public void RowColumnTable_RightAlignsCurrencyPercentAndParenthesizedNumbers() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.RightAlignedNumbers();
        style.ColumnWidthPoints = new List<double?> { 120, 100 };

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "Metric", "Amount" },
                                    new[] { "Revenue", "$1,234.50" },
                                    new[] { "Refund", "(45.20)" },
                                    new[] { "Margin", "99%" },
                                    new[] { "EU", "€1,234.50" }
                                }, style: style))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double dollarEnd = FindWordEndX(page, "$1,234.50");
        double refundEnd = FindWordEndX(page, "(45.20)");
        double percentEnd = FindWordEndX(page, "99%");
        double euroEnd = FindWordEndX(page, "€1,234.50");

        Assert.InRange(Math.Abs(refundEnd - dollarEnd), 0, 3);
        Assert.InRange(Math.Abs(percentEnd - dollarEnd), 0, 3);
        Assert.InRange(Math.Abs(euroEnd - dollarEnd), 0, 3);
    }


}
