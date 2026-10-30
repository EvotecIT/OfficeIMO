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
    public void Table_RendersConfiguredBodyColumnFills() {
        var style = PdfTableStyles.Minimal();
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.BodyColumnFills = new List<global::OfficeIMO.OfficeColor?> {
            null,
            global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.11, 0.22, 0.33)
        };

        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(320),
                PageHeight = global::OfficeIMO.OfficeLength.Points(220),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
                MarginRight = global::OfficeIMO.OfficeLength.Points(30),
                MarginTop = global::OfficeIMO.OfficeLength.Points(30),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(30)
            })
            .Table(new[] {
                new[] { "Metric", "Status" },
                new[] { "Queue", "Healthy" },
                new[] { "Latency", "Warning" }
            }, style: style)
            .ToBytes();

        string content = PdfTestContentText.ReadDecoded(bytes);
        int fillCount = content.Split(new[] { "0.11 0.22 0.33 rg" }, StringSplitOptions.None).Length - 1;

        Assert.Equal(2, fillCount);
        Assert.Contains(" re f", content);
    }

    [Fact]
    public void RowColumnTable_RendersConfiguredBodyColumnFills() {
        var style = PdfTableStyles.Minimal();
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.BodyColumnFills = new List<global::OfficeIMO.OfficeColor?> {
            null,
            global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.11, 0.22, 0.33)
        };

        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(320),
                PageHeight = global::OfficeIMO.OfficeLength.Points(220),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
                MarginRight = global::OfficeIMO.OfficeLength.Points(30),
                MarginTop = global::OfficeIMO.OfficeLength.Points(30),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(30)
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.PercentColumn(100, column =>
                                column.Table(new[] {
                                    new[] { "Metric", "Status" },
                                    new[] { "Queue", "Healthy" },
                                    new[] { "Latency", "Warning" }
                                }, style: style))))))
            .ToBytes();

        string contentStream = PdfTestContentText.ReadDecoded(bytes);
        int fillCount = contentStream.Split(new[] { "0.11 0.22 0.33 rg" }, StringSplitOptions.None).Length - 1;

        Assert.Equal(2, fillCount);
        Assert.Contains(" re f", contentStream);
    }

    [Fact]
    public void Table_DoesNotApplyBodyRowStripeFillToHeaderRows() {
        var style = PdfTableStyles.Minimal();
        style.HeaderRowCount = 2;
        style.HeaderFill = null;
        style.RowStripeFill = global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.19, 0.29, 0.39);

        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(320),
                PageHeight = global::OfficeIMO.OfficeLength.Points(220),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
                MarginRight = global::OfficeIMO.OfficeLength.Points(30),
                MarginTop = global::OfficeIMO.OfficeLength.Points(30),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(30)
            })
            .Table(new[] {
                new[] { "Group", "State" },
                new[] { "Metric", "Owner" },
                new[] { "Queue", "Healthy" },
                new[] { "Latency", "Warning" },
                new[] { "Errors", "None" }
            }, style: style)
            .ToBytes();

        string content = PdfTestContentText.ReadDecoded(bytes);
        int fillCount = content.Split(new[] { "0.19 0.29 0.39 rg" }, StringSplitOptions.None).Length - 1;

        Assert.Equal(1, fillCount);
    }

    [Fact]
    public void RowColumnTable_DoesNotApplyBodyRowStripeFillToHeaderRows() {
        var style = PdfTableStyles.Minimal();
        style.HeaderRowCount = 2;
        style.HeaderFill = null;
        style.RowStripeFill = global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.19, 0.29, 0.39);

        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(320),
                PageHeight = global::OfficeIMO.OfficeLength.Points(220),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
                MarginRight = global::OfficeIMO.OfficeLength.Points(30),
                MarginTop = global::OfficeIMO.OfficeLength.Points(30),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(30)
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.PercentColumn(100, column =>
                                column.Table(new[] {
                                    new[] { "Group", "State" },
                                    new[] { "Metric", "Owner" },
                                    new[] { "Queue", "Healthy" },
                                    new[] { "Latency", "Warning" },
                                    new[] { "Errors", "None" }
                                }, style: style))))))
            .ToBytes();

        string content = PdfTestContentText.ReadDecoded(bytes);
        int fillCount = content.Split(new[] { "0.19 0.29 0.39 rg" }, StringSplitOptions.None).Length - 1;

        Assert.Equal(1, fillCount);
    }

    [Fact]
    public void Table_StripesBodyRowsRelativeToFirstBodyRow() {
        var style = PdfTableStyles.Minimal();
        style.HeaderRowCount = 1;
        style.HeaderFill = null;
        style.RowStripeFill = global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.21, 0.31, 0.41);

        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(320),
                PageHeight = global::OfficeIMO.OfficeLength.Points(220),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
                MarginRight = global::OfficeIMO.OfficeLength.Points(30),
                MarginTop = global::OfficeIMO.OfficeLength.Points(30),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(30)
            })
            .Table(new[] {
                new[] { "Metric", "Status" },
                new[] { "Queue", "Healthy" },
                new[] { "Latency", "Warning" },
                new[] { "Errors", "None" }
            }, style: style)
            .ToBytes();

        string content = PdfTestContentText.ReadDecoded(bytes);
        int fillCount = content.Split(new[] { "0.21 0.31 0.41 rg" }, StringSplitOptions.None).Length - 1;

        Assert.Equal(1, fillCount);
    }

    [Fact]
    public void RowColumnTable_StripesBodyRowsRelativeToFirstBodyRow() {
        var style = PdfTableStyles.Minimal();
        style.HeaderRowCount = 1;
        style.HeaderFill = null;
        style.RowStripeFill = global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.21, 0.31, 0.41);

        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(320),
                PageHeight = global::OfficeIMO.OfficeLength.Points(220),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
                MarginRight = global::OfficeIMO.OfficeLength.Points(30),
                MarginTop = global::OfficeIMO.OfficeLength.Points(30),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(30)
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.PercentColumn(100, column =>
                                column.Table(new[] {
                                    new[] { "Metric", "Status" },
                                    new[] { "Queue", "Healthy" },
                                    new[] { "Latency", "Warning" },
                                    new[] { "Errors", "None" }
                                }, style: style))))))
            .ToBytes();

        string content = PdfTestContentText.ReadDecoded(bytes);
        int fillCount = content.Split(new[] { "0.21 0.31 0.41 rg" }, StringSplitOptions.None).Length - 1;

        Assert.Equal(1, fillCount);
    }

    [Fact]
    public void Table_RendersConfiguredCellFills() {
        var style = PdfTableStyles.Minimal();
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.CellFills = new Dictionary<(int Row, int Column), global::OfficeIMO.OfficeColor> {
            [(2, 1)] = global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.42, 0.18, 0.66)
        };

        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(320),
                PageHeight = global::OfficeIMO.OfficeLength.Points(220),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
                MarginRight = global::OfficeIMO.OfficeLength.Points(30),
                MarginTop = global::OfficeIMO.OfficeLength.Points(30),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(30)
            })
            .Table(new[] {
                new[] { "Metric", "Status" },
                new[] { "Queue", "Healthy" },
                new[] { "Latency", "Warning" }
            }, style: style)
            .ToBytes();

        string content = PdfTestContentText.ReadDecoded(bytes);
        int fillCount = content.Split(new[] { "0.42 0.18 0.66 rg" }, StringSplitOptions.None).Length - 1;

        Assert.Equal(1, fillCount);
        Assert.Contains(" re f", content);
    }

    [Fact]
    public void RowColumnTable_RendersConfiguredCellFills() {
        var style = PdfTableStyles.Minimal();
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.CellFills = new Dictionary<(int Row, int Column), global::OfficeIMO.OfficeColor> {
            [(2, 1)] = global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.42, 0.18, 0.66)
        };

        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(320),
                PageHeight = global::OfficeIMO.OfficeLength.Points(220),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
                MarginRight = global::OfficeIMO.OfficeLength.Points(30),
                MarginTop = global::OfficeIMO.OfficeLength.Points(30),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(30)
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.PercentColumn(100, column =>
                                column.Table(new[] {
                                    new[] { "Metric", "Status" },
                                    new[] { "Queue", "Healthy" },
                                    new[] { "Latency", "Warning" }
                                }, style: style))))))
            .ToBytes();

        string contentStream = PdfTestContentText.ReadDecoded(bytes);
        int fillCount = contentStream.Split(new[] { "0.42 0.18 0.66 rg" }, StringSplitOptions.None).Length - 1;

        Assert.Equal(1, fillCount);
        Assert.Contains(" re f", contentStream);
    }

    [Fact]
    public void Table_RendersConfiguredCellDataBarsBehindText() {
        var style = PdfTableStyles.Minimal();
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.CellDataBars = new Dictionary<(int Row, int Column), PdfCellDataBar> {
            [(1, 1)] = new PdfCellDataBar {
                Color = global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.12, 0.34, 0.56),
                Ratio = 0.5
            }
        };

        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(320),
                PageHeight = global::OfficeIMO.OfficeLength.Points(220),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
                MarginRight = global::OfficeIMO.OfficeLength.Points(30),
                MarginTop = global::OfficeIMO.OfficeLength.Points(30),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(30)
            })
            .Table(new[] {
                new[] { "Metric", "Status" },
                new[] { "Progress", "50" },
                new[] { "Done", "100" }
            }, style: style)
            .ToBytes();

        string content = PdfTestContentText.ReadDecoded(bytes);

        Assert.Contains("0.12 0.34 0.56 rg", content, StringComparison.Ordinal);
        using (PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes))) {
            Assert.Contains("50", pdf.GetPage(1).Text, StringComparison.Ordinal);
        }
        Assert.Contains(" re f", content, StringComparison.Ordinal);
    }

    [Fact]
    public void RowColumnTable_RendersConfiguredCellDataBarsBehindText() {
        var style = PdfTableStyles.Minimal();
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.CellDataBars = new Dictionary<(int Row, int Column), PdfCellDataBar> {
            [(1, 1)] = new PdfCellDataBar {
                Color = global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.12, 0.34, 0.56),
                Ratio = 0.5
            }
        };

        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(320),
                PageHeight = global::OfficeIMO.OfficeLength.Points(220),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
                MarginRight = global::OfficeIMO.OfficeLength.Points(30),
                MarginTop = global::OfficeIMO.OfficeLength.Points(30),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(30)
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.PercentColumn(100, column =>
                                column.Table(new[] {
                                    new[] { "Metric", "Status" },
                                    new[] { "Progress", "50" },
                                    new[] { "Done", "100" }
                                }, style: style))))))
            .ToBytes();

        string contentStream = PdfTestContentText.ReadDecoded(bytes);

        Assert.Contains("0.12 0.34 0.56 rg", contentStream, StringComparison.Ordinal);
        using (PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes))) {
            Assert.Contains("50", pdf.GetPage(1).Text, StringComparison.Ordinal);
        }
        Assert.Contains(" re f", contentStream, StringComparison.Ordinal);
    }

    [Fact]
    public void Table_RendersConfiguredCellIconsBeforeText() {
        var style = PdfTableStyles.Minimal();
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.CellIcons = new Dictionary<(int Row, int Column), PdfCellIcon> {
            [(1, 1)] = new PdfCellIcon {
                Kind = PdfCellIconKind.Circle,
                Color = global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.12, 0.34, 0.56),
                Size = 8
            }
        };
        style.CellPaddings = new Dictionary<(int Row, int Column), PdfCellPadding> {
            [(1, 1)] = new PdfCellPadding {
                Left = 16
            }
        };

        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(320),
                PageHeight = global::OfficeIMO.OfficeLength.Points(220),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
                MarginRight = global::OfficeIMO.OfficeLength.Points(30),
                MarginTop = global::OfficeIMO.OfficeLength.Points(30),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(30)
            })
            .Table(new[] {
                new[] { "Metric", "Status" },
                new[] { "Progress", "50" },
                new[] { "Done", "100" }
            }, style: style)
            .ToBytes();

        string content = PdfTestContentText.ReadDecoded(bytes);

        Assert.Contains("0.12 0.34 0.56 rg", content, StringComparison.Ordinal);
        Assert.Contains(" c ", content, StringComparison.Ordinal);
        using (PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes))) {
            Assert.Contains("50", pdf.GetPage(1).Text, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void RowColumnTable_RendersConfiguredCellIconsBeforeText() {
        var style = PdfTableStyles.Minimal();
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.CellIcons = new Dictionary<(int Row, int Column), PdfCellIcon> {
            [(1, 1)] = new PdfCellIcon {
                Kind = PdfCellIconKind.TriangleUp,
                Color = global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.12, 0.34, 0.56),
                Size = 8
            }
        };
        style.CellPaddings = new Dictionary<(int Row, int Column), PdfCellPadding> {
            [(1, 1)] = new PdfCellPadding {
                Left = 16
            }
        };

        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(320),
                PageHeight = global::OfficeIMO.OfficeLength.Points(220),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
                MarginRight = global::OfficeIMO.OfficeLength.Points(30),
                MarginTop = global::OfficeIMO.OfficeLength.Points(30),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(30)
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.PercentColumn(100, column =>
                                column.Table(new[] {
                                    new[] { "Metric", "Status" },
                                    new[] { "Progress", "50" },
                                    new[] { "Done", "100" }
                                }, style: style))))))
            .ToBytes();

        string contentStream = PdfTestContentText.ReadDecoded(bytes);

        Assert.Contains("0.12 0.34 0.56 rg", contentStream, StringComparison.Ordinal);
        Assert.Contains(" l ", contentStream, StringComparison.Ordinal);
        using (PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes))) {
            Assert.Contains("50", pdf.GetPage(1).Text, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void CanvasTable_RendersCellIndicatorsAfterOpaqueFillsAndBeforeText() {
        var style = PdfTableStyles.Minimal();
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.CellFills = new Dictionary<(int Row, int Column), global::OfficeIMO.OfficeColor> {
            [(0, 0)] = global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.42, 0.18, 0.66)
        };
        style.CellDataBars = new Dictionary<(int Row, int Column), PdfCellDataBar> {
            [(0, 0)] = new PdfCellDataBar {
                Color = global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.12, 0.34, 0.56),
                Ratio = 0.5
            }
        };
        style.CellIcons = new Dictionary<(int Row, int Column), PdfCellIcon> {
            [(0, 0)] = new PdfCellIcon {
                Kind = PdfCellIconKind.TriangleUp,
                Color = global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.75, 0.25, 0.1),
                Size = 8
            }
        };
        style.CellPaddings = new Dictionary<(int Row, int Column), PdfCellPadding> {
            [(0, 0)] = new PdfCellPadding {
                Left = 16
            }
        };

        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(240),
                PageHeight = global::OfficeIMO.OfficeLength.Points(160),
                CompressContentStreams = false
            })
            .Canvas(canvas => canvas.Table(new[] {
                new[] { "FixedIndicatorText" }
            }, 24, 24, 160, 60, style))
            .ToBytes();

        string content = PdfTestContentText.JoinTextArrays(string.Join("\n", GetPageContentStreams(bytes, pageNumber: 1)));
        int fill = content.IndexOf("0.42 0.18 0.66 rg", StringComparison.Ordinal);
        int dataBar = content.IndexOf("0.12 0.34 0.56 rg", StringComparison.Ordinal);
        int icon = content.IndexOf("0.75 0.25 0.1 rg", StringComparison.Ordinal);
        int text = content.IndexOf("4669786564496E64696361746F7254657874", StringComparison.Ordinal);

        using (PdfPigDocument pdf = PdfPigDocument.Open(bytes)) {
            Assert.Contains("FixedIndicatorText", pdf.GetPage(1).Text, StringComparison.Ordinal);
        }

        Assert.True(fill >= 0, "Expected the configured opaque cell fill in the page content stream.");
        Assert.True(dataBar > fill, "Expected the data bar to be painted after the opaque cell fill.");
        Assert.True(icon > dataBar, "Expected the cell icon to be painted after the opaque cell fill and data bar.");
        Assert.True(text > icon, "Expected cell text to be painted after cell indicators.");
    }


}
