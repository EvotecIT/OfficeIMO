using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentComplianceAssessmentTests {

    [Fact]
    public void TaggedFormWidgetsEmitStructureReferences() {
        byte[] pdf = PdfDocument.Create(new PdfCreateOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .TextField("Contact.Email", value: "info@example.com")
            .CheckBox("Contact.Accepted", isChecked: true)
            .ChoiceField("Contact.Country", new[] { "PL", "DE" }, value: "PL")
            .RadioButtonGroup("Contact.Approval", new[] { "Yes", "No" }, value: "Yes")
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Equal(5, CountOccurrences(content, "/Subtype /Widget"));
        Assert.Equal(5, CountOccurrences(content, "/Type /StructElem /S /Form"));
        Assert.Contains("/StructParent 0", content, StringComparison.Ordinal);
        Assert.Contains("/StructParent 1", content, StringComparison.Ordinal);
        Assert.Contains("/StructParent 2", content, StringComparison.Ordinal);
        Assert.Contains("/StructParent 3", content, StringComparison.Ordinal);
        Assert.Contains("/StructParent 4", content, StringComparison.Ordinal);
        Assert.Contains("/Type /OBJR /Obj", content, StringComparison.Ordinal);
        Assert.Contains("/ParentTree", content, StringComparison.Ordinal);
        Assert.Contains("/ParentTreeNextKey 5", content, StringComparison.Ordinal);
        Assert.Matches(@"/Nums \[0 \d+ 0 R 1 \d+ 0 R 2 \d+ 0 R 3 \d+ 0 R 4 \d+ 0 R\]", content);
    }

    [Fact]
    public void TaggedListItemsEmitLabelAndBodyStructureReferencesWithPageScopedMcids() {
        byte[] pdf = PdfDocument.Create(new PdfCreateOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .H1("Checklist")
            .Bullets(new[] { "First item", "Second item" })
            .Image(CreateMinimalRgbPng(), 24, 24, alternativeText: "Company logo")
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/StructParents 0", content, StringComparison.Ordinal);
        Assert.Contains("/H1 << /MCID 0 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/Lbl << /MCID 1 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/LBody << /MCID 2 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/Lbl << /MCID 3 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/LBody << /MCID 4 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/Figure << /Alt <436F6D70616E79206C6F676F> /MCID 5 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /H1", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /L", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /LI", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Lbl", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /LBody", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Figure", content, StringComparison.Ordinal);
        Assert.Contains("/ParentTree", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedSplitListItemKeepsContinuationBodiesUnderOneListItem() {
        string longItem = string.Join(" ", Enumerable.Repeat("Continuation text keeps one logical list item", 18));

        byte[] pdf = PdfDocument.Create(new PdfCreateOptions {
                CompressContentStreams = false,
                PageWidth = global::OfficeIMO.OfficeLength.Points(170),
                PageHeight = global::OfficeIMO.OfficeLength.Points(150),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(24),
                MarginRight = global::OfficeIMO.OfficeLength.Points(24),
                MarginTop = global::OfficeIMO.OfficeLength.Points(24),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(24),
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
            })
            .TaggedPdfCatalogMarkers()
            .Bullets(new[] { longItem })
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Equal(1, CountOccurrences(content, "/Type /StructElem /S /LI"));
        Assert.True(CountOccurrences(content, "/Type /StructElem /S /LBody") >= 2);
        Assert.Contains("/StructParents 0", content, StringComparison.Ordinal);
        Assert.Contains("/StructParents 1", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedRowColumnListItemsEmitLabelAndBodyStructureReferences() {
        byte[] pdf = PdfDocument.Create(new PdfCreateOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.PercentColumn(100, column =>
                                column.Numbered(new[] { "First item", "Second item" }))))))
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/StructParents 0", content, StringComparison.Ordinal);
        Assert.Contains("/Lbl << /MCID 0 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/LBody << /MCID 1 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/Lbl << /MCID 2 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/LBody << /MCID 3 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /L", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /LI", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Lbl", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /LBody", content, StringComparison.Ordinal);
        Assert.Contains("/ParentTree", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedTableCellsEmitStructureReferencesWithPageScopedMcids() {
        PdfTableStyle style = PdfTableStyles.Minimal();
        style.HeaderRowCount = 1;

        byte[] pdf = PdfDocument.Create(new PdfCreateOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .H1("Table")
            .Table(new[] {
                new[] { "Name", "Status" },
                new[] { "Alpha", "Ready" }
            }, style: style)
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/StructParents 0", content, StringComparison.Ordinal);
        Assert.Contains("/H1 << /MCID 0 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/TH << /MCID 1 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/TH << /MCID 2 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/TD << /MCID 3 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/TD << /MCID 4 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /H1", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Table", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /TR", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /TH", content, StringComparison.Ordinal);
        Assert.Contains("/A << /O /Table /Scope /Column >>", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /TD", content, StringComparison.Ordinal);
        Assert.Contains("/ParentTree", content, StringComparison.Ordinal);
        Assert.Contains("/Nums [0 [", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedSplitTableRowKeepsContinuationCellsUnderOneTableRow() {
        PdfTableStyle style = PdfTableStyles.Minimal();
        style.HeaderRowCount = 0;
        string longCell = string.Join(" ", Enumerable.Repeat("Continuation table row keeps one logical row", 22));

        byte[] pdf = PdfDocument.Create(new PdfCreateOptions {
                CompressContentStreams = false,
                PageWidth = global::OfficeIMO.OfficeLength.Points(170),
                PageHeight = global::OfficeIMO.OfficeLength.Points(150),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(24),
                MarginRight = global::OfficeIMO.OfficeLength.Points(24),
                MarginTop = global::OfficeIMO.OfficeLength.Points(24),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(24),
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
            })
            .TaggedPdfCatalogMarkers()
            .Table(new[] {
                new[] { longCell }
            }, style: style)
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Equal(1, CountOccurrences(content, "/Type /StructElem /S /TR"));
        Assert.True(CountOccurrences(content, "/Type /StructElem /S /TD") >= 2);
        Assert.Contains("/StructParents 0", content, StringComparison.Ordinal);
        Assert.Contains("/StructParents 1", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedLinkedTableCellWrapsTextInLinkStructure() {
        PdfTableStyle style = PdfTableStyles.Minimal();
        style.HeaderRowCount = 0;

        byte[] pdf = PdfDocument.Create(new PdfCreateOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Table(new[] {
                new[] { PdfTableCell.TextCell("Resource", linkUri: "https://officeimo.net/table", linkContents: "Linked table resource") }
            }, style: style)
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/Subtype /Link", content, StringComparison.Ordinal);
        Assert.Contains("/Link << /MCID", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Table", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /TD", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Link", content, StringComparison.Ordinal);
        Assert.Matches(@"/Type /StructElem /S /TD /P \d+ 0 R /Pg \d+ 0 R /K \[\d+ 0 R\]", content);
        Assert.Matches(@"/Type /StructElem /S /Link /P \d+ 0 R /Pg \d+ 0 R /K \[<< /Type /MCR /Pg \d+ 0 R /MCID \d+ >> << /Type /OBJR /Obj \d+ 0 R >>\]", content);
        Assert.Matches(@"/Nums \[0 \[[^\]]+\] 1 (?<link>\d+) 0 R\]", content);
    }

    [Fact]
    public void TaggedRowColumnLinkedTableCellWrapsTextInLinkStructure() {
        PdfTableStyle style = PdfTableStyles.Minimal();
        style.HeaderRowCount = 0;

        byte[] pdf = PdfDocument.Create(new PdfCreateOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.PercentColumn(100, column =>
                                column.Table(new[] {
                                    new[] { PdfTableCell.TextCell("Resource", linkUri: "https://officeimo.net/row-table", linkContents: "Linked row table resource") }
                                }, style: style))))))
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/Subtype /Link", content, StringComparison.Ordinal);
        Assert.Contains("/Link << /MCID", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Table", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /TD", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Link", content, StringComparison.Ordinal);
        Assert.Matches(@"/Type /StructElem /S /TD /P \d+ 0 R /Pg \d+ 0 R /K \[\d+ 0 R\]", content);
        Assert.Matches(@"/Type /StructElem /S /Link /P \d+ 0 R /Pg \d+ 0 R /K \[<< /Type /MCR /Pg \d+ 0 R /MCID \d+ >> << /Type /OBJR /Obj \d+ 0 R >>\]", content);
    }

    [Fact]
    public void TaggedTableDataBarsEmitArtifactMarkedContent() {
        PdfTableStyle style = PdfTableStyles.Minimal();
        style.CellDataBars = new Dictionary<(int Row, int Column), PdfCellDataBar> {
            [(1, 1)] = new PdfCellDataBar {
                Ratio = 0.75,
                Color = global::OfficeIMO.OfficeColor.FromRgb(34, 197, 94)
            }
        };

        byte[] pdf = PdfDocument.Create(new PdfCreateOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Table(new[] {
                new[] { "Name", "Progress" },
                new[] { "Alpha", "75%" }
            }, style: style)
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/Artifact << /Type /Layout >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/TD << /MCID", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /TD", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedTableCellIconsEmitArtifactMarkedContent() {
        PdfTableStyle style = PdfTableStyles.Minimal();
        style.CellIcons = new Dictionary<(int Row, int Column), PdfCellIcon> {
            [(1, 1)] = new PdfCellIcon {
                Kind = PdfCellIconKind.Circle,
                Color = global::OfficeIMO.OfficeColor.FromRgb(34, 197, 94)
            }
        };

        byte[] pdf = PdfDocument.Create(new PdfCreateOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Table(new[] {
                new[] { "Name", "Status" },
                new[] { "Alpha", "Ready" }
            }, style: style)
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/Artifact << /Type /Layout >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/TD << /MCID", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /TD", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedRowColumnTableCellIconsEmitArtifactMarkedContent() {
        PdfTableStyle style = PdfTableStyles.Minimal();
        style.CellIcons = new Dictionary<(int Row, int Column), PdfCellIcon> {
            [(1, 1)] = new PdfCellIcon {
                Kind = PdfCellIconKind.TriangleUp,
                Color = global::OfficeIMO.OfficeColor.FromRgb(14, 165, 233)
            }
        };

        byte[] pdf = PdfDocument.Create(new PdfCreateOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.PercentColumn(100, column =>
                                column.Table(new[] {
                                    new[] { "Name", "Status" },
                                    new[] { "Alpha", "Ready" }
                                }, style: style))))))
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/Artifact << /Type /Layout >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/TD << /MCID", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /TD", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedMergedTableCellsEmitSpanStructureAttributes() {
        PdfTableStyle style = PdfTableStyles.Minimal();
        style.HeaderRowCount = 1;

        byte[] pdf = PdfDocument.Create(new PdfCreateOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Table(new[] {
                new[] { PdfTableCell.Span("Group", 2) },
                new[] { PdfTableCell.Merge("Alpha", rowSpan: 2), new PdfTableCell("Ready") },
                new[] { new PdfTableCell("Done") }
            }, style: style)
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/TH << /MCID 0 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/TD << /MCID 1 >> BDC", content, StringComparison.Ordinal);
        var elements = PdfReadDocument.Open(pdf).RawStructure().Objects.Where(item =>
            item.Value.Kind == PdfRawValueKind.Dictionary && item.Value.Entries.TryGetValue("Type", out var type) &&
            type.Text == "StructElem").ToArray();
        var header = Assert.Single(elements, item => item.Value.Entries["S"].Text == "TH");
        Assert.Equal("Column", header.Value.Entries["A"].Entries["Scope"].Text);
        Assert.Equal(2D, header.Value.Entries["A"].Entries["ColSpan"].Number);
        var merged = Assert.Single(elements, item => item.Value.Entries["S"].Text == "TD" &&
            item.Value.Entries["A"].Entries.ContainsKey("RowSpan"));
        Assert.Equal(2D, merged.Value.Entries["A"].Entries["RowSpan"].Number);
        Assert.Equal(header.Value.Entries["ID"].Text,
            Assert.Single(merged.Value.Entries["A"].Entries["Headers"].Items).Text);
    }

    [Fact]
    public void TaggedTableCaptionEmitsCaptionStructureReferences() {
        PdfTableStyle style = PdfTableStyles.Minimal();
        style.HeaderRowCount = 1;
        style.Caption = "Table 1. Status signals";

        byte[] pdf = PdfDocument.Create(new PdfCreateOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Table(new[] {
                new[] { "Name", "Status" },
                new[] { "Alpha", "Ready" }
            }, style: style)
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/Caption << /MCID 0 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/TH << /MCID 1 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Table", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Caption", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /TR", content, StringComparison.Ordinal);
        Assert.Contains("/ParentTree", content, StringComparison.Ordinal);
    }

    [Fact]
    public void TaggedTableAlternativeTextEmitsTableStructureAltText() {
        PdfTableStyle style = PdfTableStyles.Minimal();
        style.HeaderRowCount = 1;
        style.AlternativeText = "Quarterly status summary";

        byte[] pdf = PdfDocument.Create(new PdfCreateOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Table(new[] {
                new[] { "Name", "Status" },
                new[] { "Alpha", "Ready" }
            }, style: style)
            .ToBytes();

        PdfDocumentInfo info = PdfInspector.Inspect(pdf);
        PdfTaggedContentInfo tagged = Assert.IsType<PdfTaggedContentInfo>(info.TaggedContent);
        PdfStructureElementInfo table = Assert.Single(tagged.StructureElements, element => element.StructureType == "Table");

        Assert.Equal("Quarterly status summary", table.AlternateText);
    }

    [Fact]
    public void TaggedRowColumnTableCaptionEmitsCaptionStructureReferences() {
        PdfTableStyle style = PdfTableStyles.Minimal();
        style.HeaderRowCount = 1;
        style.Caption = "Table 1. Column signals";

        byte[] pdf = PdfDocument.Create(new PdfCreateOptions {
                CompressContentStreams = false
            })
            .TaggedPdfCatalogMarkers()
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.PercentColumn(100, column =>
                                column.Table(new[] {
                                    new[] { "Name", "Status" },
                                    new[] { "Alpha", "Ready" }
                                }, style: style))))))
            .ToBytes();

        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/Caption << /MCID 0 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/TH << /MCID 1 >> BDC", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Table", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /Caption", content, StringComparison.Ordinal);
        Assert.Contains("/Type /StructElem /S /TR", content, StringComparison.Ordinal);
        Assert.Contains("/ParentTree", content, StringComparison.Ordinal);
    }

}
