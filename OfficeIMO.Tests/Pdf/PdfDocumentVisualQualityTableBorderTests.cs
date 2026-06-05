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
    public void Table_RendersConfiguredCellBorders() {
        var style = TableStyles.Minimal();
        style.BorderColor = null;
        style.CellBorders = new Dictionary<(int Row, int Column), PdfCellBorder> {
            [(2, 1)] = new PdfCellBorder {
                Color = new PdfColor(0.12, 0.34, 0.56),
                Width = 1.7
            }
        };

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Table(new[] {
                new[] { "Metric", "Status" },
                new[] { "Queue", "Healthy" },
                new[] { "Latency", "Warning" }
            }, style: style)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        int borderColorCount = content.Split(new[] { "0.12 0.34 0.56 RG" }, StringSplitOptions.None).Length - 1;

        Assert.Equal(1, borderColorCount);
        Assert.Contains("1.7 w", content);
        Assert.Contains(" re S", content);
    }

    [Fact]
    public void Table_RendersConfiguredSideSpecificCellBorders() {
        var style = TableStyles.Minimal();
        style.BorderColor = null;
        style.CellBorders = new Dictionary<(int Row, int Column), PdfCellBorder> {
            [(1, 0)] = new PdfCellBorder {
                Color = null,
                Width = 0,
                Top = true,
                Right = false,
                Bottom = false,
                Left = true,
                TopBorder = new PdfCellBorderSide {
                    Color = PdfColor.FromRgb(255, 0, 0),
                    Width = 2
                },
                LeftBorder = new PdfCellBorderSide {
                    Color = PdfColor.FromRgb(0, 0, 255),
                    Width = 1.5
                }
            }
        };

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Table(new[] {
                new[] { "Metric", "Status" },
                new[] { "Queue", "Healthy" }
            }, style: style)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("1 0 0 RG", content);
        Assert.Contains("2 w", content);
        Assert.Contains("0 0 1 RG", content);
        Assert.Contains("1.5 w", content);
    }

    [Fact]
    public void Table_RendersConfiguredDashedCellBorders() {
        var style = TableStyles.Minimal();
        style.BorderColor = null;
        style.CellBorders = new Dictionary<(int Row, int Column), PdfCellBorder> {
            [(1, 0)] = new PdfCellBorder {
                Color = PdfColor.FromRgb(18, 52, 86),
                Width = 1,
                DashStyle = OfficeStrokeDashStyle.Dash
            },
            [(1, 1)] = new PdfCellBorder {
                Color = null,
                TopBorder = new PdfCellBorderSide {
                    Color = PdfColor.FromRgb(120, 80, 40),
                    Width = 1.5,
                    DashStyle = OfficeStrokeDashStyle.Dot
                },
                BottomBorder = new PdfCellBorderSide {
                    Color = PdfColor.FromRgb(40, 120, 80),
                    Width = 1.25,
                    DashStyle = OfficeStrokeDashStyle.DashDot
                }
            }
        };

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Table(new[] {
                new[] { "Metric", "Status" },
                new[] { "Queue", "Healthy" }
            }, style: style)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("[3 1.5] 0 d", content, StringComparison.Ordinal);
        Assert.Contains("[1.5 2.25] 0 d", content, StringComparison.Ordinal);
        Assert.Contains("[3.75 1.875 1.25 1.875] 0 d", content, StringComparison.Ordinal);
        Assert.Contains("1 J", content, StringComparison.Ordinal);
        Assert.True(content.Contains(" m ", StringComparison.Ordinal) && content.Contains(" l S", StringComparison.Ordinal), "Expected diagonal cell borders to emit line segments instead of only rectangle borders.");
    }

    [Fact]
    public void Table_RendersConfiguredDoubleAndDiagonalCellBorders() {
        var style = TableStyles.Minimal();
        style.BorderColor = null;
        style.CellBorders = new Dictionary<(int Row, int Column), PdfCellBorder> {
            [(1, 0)] = new PdfCellBorder {
                Color = PdfColor.FromRgb(68, 85, 102),
                Width = 1,
                LineStyle = PdfCellBorderLineStyle.TwoLine,
                DiagonalUp = true,
                DiagonalDown = true
            },
            [(1, 1)] = new PdfCellBorder {
                Color = null,
                TopBorder = new PdfCellBorderSide {
                    Color = PdfColor.FromRgb(120, 80, 40),
                    Width = 1.25,
                    LineStyle = PdfCellBorderLineStyle.TwoLine
                },
                DiagonalDown = true,
                DiagonalDownBorder = new PdfCellBorderSide {
                    Color = PdfColor.FromRgb(40, 120, 80),
                    Width = 0.75,
                    DashStyle = OfficeStrokeDashStyle.Dash,
                    LineStyle = PdfCellBorderLineStyle.TwoLine
                }
            }
        };

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Table(new[] {
                new[] { "Metric", "Status" },
                new[] { "Queue", "Healthy" }
            }, style: style)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("0.267 0.333 0.4 RG", content, StringComparison.Ordinal);
        Assert.Contains("0.157 0.471 0.314 RG", content, StringComparison.Ordinal);
        Assert.Contains("[2.25 1.125] 0 d", content, StringComparison.Ordinal);
        Assert.True(content.Split(new[] { " S" }, StringSplitOptions.None).Length - 1 >= 10, "Expected double and diagonal cell borders to emit multiple stroked lines.");
        Assert.True(content.Contains(" m ", StringComparison.Ordinal) && content.Contains(" l S", StringComparison.Ordinal), "Expected diagonal cell borders to emit line segments instead of only rectangle borders.");
    }

    [Fact]
    public void Table_RendersConfiguredCellBordersAfterCellText() {
        var style = TableStyles.Minimal();
        style.BorderColor = null;
        style.CellBorders = new Dictionary<(int Row, int Column), PdfCellBorder> {
            [(0, 0)] = new PdfCellBorder {
                Color = PdfColor.FromRgb(255, 0, 0),
                Width = 3,
                DiagonalDown = true
            }
        };

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 160,
                MarginLeft = 24,
                MarginRight = 24,
                MarginTop = 24,
                MarginBottom = 24
            })
            .Table(new[] {
                new[] { "BorderText" }
            }, style: style)
            .ToBytes();

        string content = string.Join("\n", GetPageContentStreams(bytes, pageNumber: 1));
        int text = content.IndexOf("<426F7264657254657874>", StringComparison.Ordinal);
        int borderColor = content.LastIndexOf("1 0 0 RG", StringComparison.Ordinal);
        int stroke = borderColor < 0 ? -1 : content.IndexOf(" S", borderColor, StringComparison.Ordinal);

        using (PdfPigDocument pdf = PdfPigDocument.Open(bytes)) {
            Assert.Contains("BorderText", pdf.GetPage(1).Text, StringComparison.Ordinal);
        }

        Assert.True(text >= 0, "Expected encoded table cell text in the page content stream.");
        Assert.True(borderColor > text, "Expected configured cell borders to be painted after table cell text.");
        Assert.True(stroke > borderColor, "Expected configured cell borders to emit a stroke after their color.");
    }

    [Fact]
    public void RowColumnTable_RendersConfiguredCellBorders() {
        var style = TableStyles.Minimal();
        style.BorderColor = null;
        style.CellBorders = new Dictionary<(int Row, int Column), PdfCellBorder> {
            [(2, 1)] = new PdfCellBorder {
                Color = new PdfColor(0.12, 0.34, 0.56),
                Width = 1.7
            }
        };

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "Metric", "Status" },
                                    new[] { "Queue", "Healthy" },
                                    new[] { "Latency", "Warning" }
                                }, style: style))))))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        int borderColorCount = content.Split(new[] { "0.12 0.34 0.56 RG" }, StringSplitOptions.None).Length - 1;

        Assert.Equal(1, borderColorCount);
        Assert.Contains("1.7 w", content);
        Assert.Contains(" re S", content);
    }

    [Fact]
    public void Table_RendersConfiguredCellBorderSides() {
        var style = TableStyles.Minimal();
        style.BorderColor = null;
        style.CellBorders = new Dictionary<(int Row, int Column), PdfCellBorder> {
            [(2, 1)] = new PdfCellBorder {
                Color = new PdfColor(0.2, 0.3, 0.4),
                Width = 2.2,
                Right = false,
                Bottom = false,
                Left = false
            }
        };

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Table(new[] {
                new[] { "Metric", "Status" },
                new[] { "Queue", "Healthy" },
                new[] { "Total", "Warning" }
            }, style: style)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        int borderColorCount = content.Split(new[] { "0.2 0.3 0.4 RG" }, StringSplitOptions.None).Length - 1;

        Assert.Equal(1, borderColorCount);
        Assert.Contains("2.2 w", content);
        Assert.Contains(" l S", content);
        Assert.DoesNotContain(" re S", content);
    }

    [Fact]
    public void Table_RendersConfiguredRowSeparatorsWithoutCellBorderDictionary() {
        var style = TableStyles.Minimal();
        style.BorderColor = null;
        style.RowSeparatorColor = new PdfColor(0.12, 0.34, 0.56);
        style.RowSeparatorWidth = 0.6;
        style.HeaderSeparatorColor = new PdfColor(0.7, 0.2, 0.1);
        style.HeaderSeparatorWidth = 1.1;

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Table(new[] {
                new[] { "Metric", "Status" },
                new[] { "Queue", "Healthy" },
                new[] { "Latency", "Warning" }
            }, style: style)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        int bodySeparatorCount = content.Split(new[] { "0.12 0.34 0.56 RG" }, StringSplitOptions.None).Length - 1;
        int headerSeparatorCount = content.Split(new[] { "0.7 0.2 0.1 RG" }, StringSplitOptions.None).Length - 1;

        Assert.Equal(2, bodySeparatorCount);
        Assert.Equal(1, headerSeparatorCount);
        Assert.Contains("0.6 w", content);
        Assert.Contains("1.1 w", content);
        Assert.Contains(" l S", content);
        Assert.DoesNotContain(" re S", content);
    }

    [Fact]
    public void Table_DrawsHeaderRulesBeforeHeaderText() {
        var style = TableStyles.Minimal();
        style.HeaderFill = PdfColor.FromRgb(15, 23, 42);
        style.HeaderTextColor = PdfColor.White;
        style.BorderColor = PdfColor.FromRgb(238, 51, 68);
        style.BorderWidth = 1.1;
        style.HeaderSeparatorColor = PdfColor.FromRgb(238, 51, 68);
        style.HeaderSeparatorWidth = 1.1;

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 160,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                CompressContentStreams = false
            })
            .Table(new[] {
                new[] { "HeaderGlyphClearance", "Value" }
            }, style: style)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        int borderColor = content.IndexOf("0.933 0.2 0.267 RG", StringComparison.Ordinal);
        int firstTextObject = content.IndexOf("BT\n", StringComparison.Ordinal);

        Assert.True(borderColor >= 0, "Expected the configured header rule color in the content stream.");
        Assert.True(firstTextObject >= 0, "Expected table header text in the content stream.");
        Assert.True(borderColor < firstTextObject, "Table header rules should be painted before text so strokes cannot cut through glyphs.");
    }

    [Fact]
    public void RowColumnTable_RendersConfiguredRowSeparatorsWithoutCellBorderDictionary() {
        var style = TableStyles.Minimal();
        style.BorderColor = null;
        style.RowSeparatorColor = new PdfColor(0.12, 0.34, 0.56);
        style.RowSeparatorWidth = 0.6;
        style.HeaderSeparatorColor = new PdfColor(0.7, 0.2, 0.1);
        style.HeaderSeparatorWidth = 1.1;

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "Metric", "Status" },
                                    new[] { "Queue", "Healthy" },
                                    new[] { "Latency", "Warning" }
                                }, style: style))))))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        int bodySeparatorCount = content.Split(new[] { "0.12 0.34 0.56 RG" }, StringSplitOptions.None).Length - 1;
        int headerSeparatorCount = content.Split(new[] { "0.7 0.2 0.1 RG" }, StringSplitOptions.None).Length - 1;

        Assert.Equal(2, bodySeparatorCount);
        Assert.Equal(1, headerSeparatorCount);
        Assert.Contains("0.6 w", content);
        Assert.Contains("1.1 w", content);
        Assert.Contains(" l S", content);
        Assert.DoesNotContain(" re S", content);
    }

    [Fact]
    public void Table_RendersConfiguredFooterSeparatorAboveFooterRows() {
        var style = TableStyles.Minimal();
        style.BorderColor = null;
        style.RowSeparatorColor = new PdfColor(0.12, 0.34, 0.56);
        style.RowSeparatorWidth = 0.6;
        style.FooterRowCount = 1;
        style.FooterSeparatorColor = new PdfColor(0.2, 0.7, 0.3);
        style.FooterSeparatorWidth = 1.3;

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Table(new[] {
                new[] { "Metric", "Status" },
                new[] { "Queue", "Healthy" },
                new[] { "Latency", "Warning" },
                new[] { "Total", "Ready" }
            }, style: style)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        int footerSeparatorCount = content.Split(new[] { "0.2 0.7 0.3 RG" }, StringSplitOptions.None).Length - 1;

        Assert.Equal(1, footerSeparatorCount);
        Assert.Contains("1.3 w", content);
        Assert.Contains(" l S", content);
        Assert.DoesNotContain(" re S", content);
    }

    [Fact]
    public void RowColumnTable_RendersConfiguredFooterSeparatorAboveFooterRows() {
        var style = TableStyles.Minimal();
        style.BorderColor = null;
        style.RowSeparatorColor = new PdfColor(0.12, 0.34, 0.56);
        style.RowSeparatorWidth = 0.6;
        style.FooterRowCount = 1;
        style.FooterSeparatorColor = new PdfColor(0.2, 0.7, 0.3);
        style.FooterSeparatorWidth = 1.3;

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "Metric", "Status" },
                                    new[] { "Queue", "Healthy" },
                                    new[] { "Latency", "Warning" },
                                    new[] { "Total", "Ready" }
                                }, style: style))))))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        int footerSeparatorCount = content.Split(new[] { "0.2 0.7 0.3 RG" }, StringSplitOptions.None).Length - 1;

        Assert.Equal(1, footerSeparatorCount);
        Assert.Contains("1.3 w", content);
        Assert.Contains(" l S", content);
        Assert.DoesNotContain(" re S", content);
    }


}
