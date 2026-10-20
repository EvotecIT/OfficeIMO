using System.Globalization;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTaggedContentReadBoundsTests {
    [Fact]
    public void ReadingSkipsContainerCyclesAndRetainsRepeatedSiblingContent() {
        byte[] source = BuildPdf(
            "[8 0 R 6 0 R]",
            "[9 0 R 10 0 R 11 0 R 11 0 R]",
            "<< /Nums [0 [6 0 R]] /Kids [7 0 R] >>",
            new[] {
                "[8 0 R 6 0 R]",
                "[0 9 0 R 1]",
                "<< /Pg 3 0 R /K [10 0 R 2] >>",
                "[3]"
            });

        PdfTaggedContentInfo tags = Assert.IsType<PdfTaggedContentInfo>(PdfReadDocument.Open(source).TaggedContent);

        Assert.Equal(new[] { 6 }, tags.RootElementObjectNumbers);
        Assert.Equal(new[] { 0 }, tags.ParentTreeStructParentIndexes);
        PdfStructureElementInfo element = Assert.Single(tags.StructureElements);
        Assert.Equal(new[] { 0, 1, 2, 3, 3 }, element.MarkedContentReferences.Select(reference => reference.MarkedContentId));
        Assert.All(element.MarkedContentReferences, reference => Assert.Equal(3, reference.PageObjectNumber));
    }

    [Theory]
    [InlineData("array")]
    [InlineData("dictionary")]
    [InlineData("parent-tree")]
    public void ReadingBoundsIndirectTaggedTraversalWithoutTruncatingSupportedDepth(string traversal) {
        byte[] supported = BuildDepthPdf(traversal, 256);
        byte[] excessive = BuildDepthPdf(traversal, 257);

        PdfTaggedContentInfo tags = Assert.IsType<PdfTaggedContentInfo>(PdfReadDocument.Open(supported).TaggedContent);

        Assert.Equal(new[] { 6 }, tags.RootElementObjectNumbers);
        if (traversal == "parent-tree") {
            Assert.Equal(new[] { 0 }, tags.ParentTreeStructParentIndexes);
        } else {
            PdfMarkedContentReference reference = Assert.Single(Assert.Single(tags.StructureElements).MarkedContentReferences);
            Assert.Equal(0, reference.MarkedContentId);
            Assert.Equal(3, reference.PageObjectNumber);
        }
        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => PdfReadDocument.Open(excessive));
        Assert.Contains("structure nesting limit", exception.Message, StringComparison.Ordinal);
    }

    private static byte[] BuildDepthPdf(string traversal, int depth) {
        var extraObjects = new List<string>();
        if (traversal == "parent-tree") {
            for (int index = 1; index <= depth; index++) {
                extraObjects.Add(index == depth
                    ? "<< /Nums [0 [6 0 R]] >>"
                    : "<< /Kids [" + Reference(index + 8) + "] >>");
            }
            return BuildPdf("[6 0 R]", "0", "<< /Kids [8 0 R] >>", extraObjects);
        }

        for (int index = 0; index < depth; index++) {
            string child = index + 1 == depth ? "0" : Reference(index + 9);
            extraObjects.Add(traversal == "array" ? "[" + child + "]" : "<< /K " + child + " >>");
        }
        return BuildPdf("[6 0 R]", "8 0 R", "<< /Nums [0 [6 0 R]] >>", extraObjects);
    }

    private static string Reference(int number) => number.ToString(CultureInfo.InvariantCulture) + " 0 R";

    private static byte[] BuildPdf(string rootChildren, string elementChildren, string parentTree, IEnumerable<string> extraObjects) {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R /MarkInfo << /Marked true >> /StructTreeRoot 5 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /StructParents 0 >>",
            "<< /Length 0 >>\nstream\n\nendstream",
            "<< /Type /StructTreeRoot /K " + rootChildren + " /ParentTree 7 0 R >>",
            "<< /Type /StructElem /S /P /P 5 0 R /Pg 3 0 R /K " + elementChildren + " >>",
            parentTree
        };
        objects.AddRange(extraObjects);
        var pdf = new StringBuilder("%PDF-1.7\n");
        var offsets = new List<int>();
        for (int index = 0; index < objects.Count; index++) {
            offsets.Add(pdf.Length);
            pdf.Append(index + 1).Append(" 0 obj\n").Append(objects[index]).Append("\nendobj\n");
        }
        int xref = pdf.Length;
        pdf.Append("xref\n0 ").Append(objects.Count + 1).Append("\n0000000000 65535 f \n");
        foreach (int offset in offsets) pdf.Append(offset.ToString("D10", CultureInfo.InvariantCulture)).Append(" 00000 n \n");
        pdf.Append("trailer << /Root 1 0 R /Size ").Append(objects.Count + 1).Append(" >>\nstartxref\n").Append(xref).Append("\n%%EOF\n");
        return Encoding.ASCII.GetBytes(pdf.ToString());
    }
}
