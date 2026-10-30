using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordBorderTabPhysicalLengthTests {
    [Fact]
    public void BorderLengthsPersistEighthPointsWholePointsZeroAndUnsetAttributes() {
        using WordDocument document = WordDocument.Create();
        WordBorders page = document.Sections[0].Borders;
        WordParagraph paragraph = document.AddParagraph("Border lengths");
        WordParagraphBorders borders = paragraph.Borders;
        Assert.Null(page.LeftSize);
        Assert.Null(page.LeftSpace);
        Assert.Null(borders.LeftSize);
        Assert.Null(borders.LeftSpace);

        page.LeftStyle = WordBorderStyle.Single;
        page.LeftSize = OfficeLength.Points(1.1);
        page.LeftSpace = OfficeLength.Points(2.5);
        page.RightSize = OfficeLength.Zero;
        page.RightSpace = OfficeLength.Zero;
        page.TopSize = OfficeLength.Points(1);
        page.TopSpace = OfficeLength.Points(1);
        page.TopSize = null;
        page.TopSpace = null;
        page.BottomStyle = WordBorderStyle.Single;
        borders.SetBorder(WordParagraphBorderType.Left, WordBorderStyle.Single, OfficeColor.Red,
            OfficeLength.Points(1.1), false);
        borders.LeftSpace = OfficeLength.Points(2.5);
        borders.RightSize = OfficeLength.Zero;
        borders.RightSpace = OfficeLength.Zero;
        borders.TopSize = OfficeLength.Points(1);
        borders.TopSpace = OfficeLength.Points(1);
        borders.TopSize = null;
        borders.TopSpace = null;
        borders.BottomStyle = WordBorderStyle.Single;

        using var bytes = new MemoryStream(document.ToBytes());
        using (WordprocessingDocument saved = WordprocessingDocument.Open(bytes, false)) {
            Body body = saved.MainDocumentPart!.Document.Body!;
            PageBorders nativePage = body.Elements<SectionProperties>().Single().GetFirstChild<PageBorders>()!;
            ParagraphBorders nativeParagraph = body.Elements<Paragraph>().Single().ParagraphProperties!.ParagraphBorders!;
            CheckNative(nativePage.LeftBorder!, nativePage.RightBorder!, nativePage.TopBorder!, nativePage.BottomBorder!);
            CheckNative(nativeParagraph.LeftBorder!, nativeParagraph.RightBorder!, nativeParagraph.TopBorder!, nativeParagraph.BottomBorder!);
        }
        bytes.Position = 0;
        using WordDocument reopened = WordDocument.Load(bytes);
        Assert.Equal(OfficeLength.Points(1.125), reopened.Borders.LeftSize);
        Assert.Equal(OfficeLength.Points(3), reopened.Borders.LeftSpace);
        Assert.Equal(OfficeLength.Zero, reopened.Borders.RightSize);
        Assert.Equal(OfficeLength.Zero, reopened.Borders.RightSpace);
        Assert.Null(reopened.Borders.TopSize);
        Assert.Null(reopened.Borders.TopSpace);
        Assert.Null(reopened.Borders.BottomSize);
        Assert.Null(reopened.Borders.BottomSpace);
        WordParagraphBorders restored = reopened.Paragraphs[0].Borders;
        Assert.Equal(OfficeLength.Points(1.125), restored.LeftSize);
        Assert.Equal(OfficeLength.Points(3), restored.LeftSpace);
        Assert.Equal(OfficeLength.Zero, restored.RightSize);
        Assert.Equal(OfficeLength.Zero, restored.RightSpace);
        Assert.Null(restored.TopSize);
        Assert.Null(restored.TopSpace);
        Assert.Null(restored.BottomSize);
        Assert.Null(restored.BottomSpace);
        WordParagraphBorderInfo snapshot = Assert.Single(reopened.Inspect().Sections)
            .Elements.OfType<WordParagraphInfo>().Single().LeftBorder!;
        Assert.Equal(restored.LeftSize, snapshot.Size);
        Assert.Equal(restored.LeftSpace, snapshot.Space);

        static void CheckNative(BorderType left, BorderType right, BorderType cleared, BorderType unset) {
            Assert.Equal(9U, left.Size!.Value);
            Assert.Equal(3U, left.Space!.Value);
            Assert.Equal(0U, right.Size!.Value);
            Assert.Equal(0U, right.Space!.Value);
            Assert.Null(cleared.Size);
            Assert.Null(cleared.Space);
            Assert.Null(unset.Size);
            Assert.Null(unset.Space);
        }
    }

    [Fact]
    public void InvalidBorderLengthsDoNotCreateOrChangeFormatting() {
        using WordDocument document = WordDocument.Create();
        WordSection section = document.Sections[0];
        WordParagraph paragraph = document.AddParagraph("No borders");
        WordBorders page = section.Borders;
        WordParagraphBorders borders = paragraph.Borders;
        Action<OfficeLength?>[] setters = {
            value => page.LeftSize = value, value => page.RightSize = value,
            value => page.TopSize = value, value => page.BottomSize = value,
            value => page.LeftSpace = value, value => page.RightSpace = value,
            value => page.TopSpace = value, value => page.BottomSpace = value,
            value => borders.LeftSize = value, value => borders.RightSize = value,
            value => borders.TopSize = value, value => borders.BottomSize = value,
            value => borders.LeftSpace = value, value => borders.RightSpace = value,
            value => borders.TopSpace = value, value => borders.BottomSpace = value
        };
        // Exercise both absent containers and existing authored values.
        for (int pass = 0; pass < 2; pass++) {
            string sectionBefore = section._sectionProperties.OuterXml;
            string paragraphBefore = paragraph._paragraph.OuterXml;
            foreach (Action<OfficeLength?> set in setters) {
                Assert.Throws<ArgumentOutOfRangeException>(() => set(OfficeLength.Emus(-1)));
                Assert.Throws<OverflowException>(() => set(OfficeLength.Emus(long.MaxValue)));
                Assert.Equal(sectionBefore, section._sectionProperties.OuterXml);
                Assert.Equal(paragraphBefore, paragraph._paragraph.OuterXml);
            }
            foreach (WordParagraphBorderType side in new[] { WordParagraphBorderType.Left,
                WordParagraphBorderType.Right, WordParagraphBorderType.Top, WordParagraphBorderType.Bottom }) {
                Assert.Throws<ArgumentOutOfRangeException>(() => borders.SetBorder(side,
                    WordBorderStyle.Double, OfficeColor.Blue, OfficeLength.Emus(-1), true));
                Assert.Throws<OverflowException>(() => borders.SetBorder(side,
                    WordBorderStyle.Double, OfficeColor.Blue, OfficeLength.Emus(long.MaxValue), true));
                Assert.Equal(paragraphBefore, paragraph._paragraph.OuterXml);
            }
            foreach (Action<OfficeLength?> set in setters) set(OfficeLength.Points(1));
        }
    }

    [Fact]
    public void SignedTabsRoundAndPersistClearStopsWithoutMutatingOnInvalidInput() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("Signed tabs");
        string before = paragraph._paragraph.OuterXml;
        Assert.Equal(OfficeLength.Zero, new WordTabStop(paragraph).Position);
        Assert.Throws<OverflowException>(() => paragraph.AddTabStop(OfficeLength.Emus(long.MaxValue)));
        Assert.Throws<OverflowException>(() => paragraph.AddTabStop(OfficeLength.Emus(long.MinValue)));
        Assert.Equal(before, paragraph._paragraph.OuterXml);
        WordTabStop negative = paragraph.AddTabStop(OfficeLength.Points(-3.03), WordTabAlignment.Right, WordTabLeader.Dot);
        negative.Position = OfficeLength.Points(-3.08);
        paragraph.AddTabStop(OfficeLength.Points(3.03), WordTabAlignment.Decimal);
        paragraph.AddTabStop(OfficeLength.Zero, WordTabAlignment.Clear);
        paragraph.AddTabStop(OfficeLength.Emus((long)int.MinValue * 635));
        paragraph.AddTabStop(OfficeLength.Emus((long)int.MaxValue * 635));
        before = paragraph._paragraph.OuterXml;
        Assert.Throws<OverflowException>(() => negative.Position = OfficeLength.Emus(long.MaxValue));
        Assert.Throws<OverflowException>(() => negative.Position = OfficeLength.Emus(long.MinValue));
        Assert.Throws<OverflowException>(() => paragraph.AddTabStop(OfficeLength.Emus(long.MaxValue)));
        Assert.Equal(before, paragraph._paragraph.OuterXml);
        using var bytes = new MemoryStream(document.ToBytes());
        using (WordprocessingDocument saved = WordprocessingDocument.Open(bytes, false)) {
            TabStop[] tabs = saved.MainDocumentPart!.Document.Body!.Elements<Paragraph>().Single()
                .ParagraphProperties!.Tabs!.Elements<TabStop>().ToArray();
            Assert.Equal(new[] { -62, 61, 0, int.MinValue, int.MaxValue }, tabs.Select(tab => tab.Position!.Value).ToArray());
            Assert.Equal(TabStopValues.Clear, tabs[2].Val!.Value);
            Assert.Equal(TabStopLeaderCharValues.Dot, tabs[0].Leader!.Value);
        }
        bytes.Position = 0;
        using WordDocument reopened = WordDocument.Load(bytes);
        WordParagraph restored = reopened.Paragraphs[0];
        Assert.Equal(OfficeLength.Points(-3.1), restored.TabStops[0].Position);
        Assert.Equal(OfficeLength.Points(3.05), restored.TabStops[1].Position);
        Assert.Equal(WordTabAlignment.Clear, restored.TabStops[2].Alignment);
        Assert.Equal(OfficeLength.Zero, restored.TabStops[2].Position);
        Assert.Equal(OfficeLength.Emus((long)int.MinValue * 635), restored.TabStops[3].Position);
        Assert.Equal(OfficeLength.Emus((long)int.MaxValue * 635), restored.TabStops[4].Position);
        restored.ClearTabStops();
        using WordDocument cleared = WordDocument.Load(new MemoryStream(reopened.ToBytes()));
        Assert.Empty(cleared.Paragraphs[0].TabStops);
    }

    [Fact]
    public void BorderStorageLimitsRemainAvailableWithoutNarrowingTheNativeRange() {
        using WordDocument document = WordDocument.Create();
        WordBorders borders = document.Borders;
        borders.LeftSize = OfficeLength.Points(uint.MaxValue / 8D);
        borders.LeftSpace = OfficeLength.Emus((long)uint.MaxValue * 12700);
        using var bytes = new MemoryStream(document.ToBytes());
        using WordprocessingDocument saved = WordprocessingDocument.Open(bytes, false);
        LeftBorder border = saved.MainDocumentPart!.Document.Body!.Elements<SectionProperties>().Single()
            .GetFirstChild<PageBorders>()!.LeftBorder!;
        Assert.Equal(uint.MaxValue, border.Size!.Value);
        Assert.Equal(uint.MaxValue, border.Space!.Value);
    }
}
