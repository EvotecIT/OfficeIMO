using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointLayoutTableTests {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ContinuationsPreserveEveryDataRowAndTheRequestedHeaderPolicy(bool includeHeaders) {
        using var document = PowerPointDocument.Create(new PowerPointCreateOptions { Theme = OfficeTheme.Modern });
        var table = new PowerPointLayoutTable(new[] { "Service", "Count" },
            Enumerable.Range(1, 7).Select(index => (IReadOnlyList<string>)new[] { "Service " + index, index.ToString() }).ToArray()) {
            IncludeHeaders = includeHeaders, MaximumRows = 3
        };
        var options = PowerPointCompositionOptions.FromDesign(PowerPointDeckDesign.FromBrand("#2563EB", "table-content"));
        options.ApplyTheme = false;
        document.Compose(new PowerPointDeckPlan().AddContent("Services", new[] { table }), options);
        using var saved = PowerPointDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Equal(3, saved.Slides.Count);
        var tables = saved.Slides.SelectMany(slide => slide.Tables).ToArray();
        Assert.Equal(3, tables.Length);
        var values = new List<string>();
        foreach (var native in tables) {
            Assert.NotNull(native.AltText);
            Assert.Contains("Service", native.AltText);
            Assert.Contains("Count", native.AltText);
            if (includeHeaders) Assert.Equal("Service", native.GetCell(0, 0).Text);
            for (int row = includeHeaders ? 1 : 0; row < native.Rows; row++) values.Add(native.GetCell(row, 0).Text);
        }
        Assert.Equal(Enumerable.Range(1, 7).Select(index => "Service " + index), values);
        Assert.Empty(saved.ValidateDocument());
        Assert.DoesNotContain(saved.InspectAccessibility().Findings,
            finding => finding.Code == "accessibility-missing-alternative-text");
    }

    [Fact]
    public void EveryAuthoredLineBreakInACellIsWrittenAsADrawingMlBreak() {
        using var document = PowerPointDocument.Create();
        var table = new PowerPointLayoutTable(new[] { "Service\vOwner", "Detail" }, new IReadOnlyList<string>[] {
            new[] { "one\ntwo", "unix" }, new[] { "one\r\ntwo", "windows" }, new[] { "one\vtwo", "soft" }
        });
        var options = PowerPointCompositionOptions.FromDesign(PowerPointDeckDesign.FromBrand("#2563EB", "table-breaks"));
        options.ApplyTheme = false;
        document.Compose(new PowerPointDeckPlan().AddContent("Services", new[] { table }), options);
        using var saved = PowerPointDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Empty(saved.ValidateDocument());
        var native = saved.Slides.SelectMany(slide => slide.Tables).Single();
        Assert.Equal(new[] { "Service", "Owner" }, native.GetCell(0, 0).Paragraphs[0].InlineNodes
            .Where(node => node.Kind == PowerPointParagraphInlineKind.Run).Select(node => node.Text));
        for (int row = 1; row <= 3; row++) {
            var inline = native.GetCell(row, 0).Paragraphs[0].InlineNodes;
            Assert.Equal(new[] { PowerPointParagraphInlineKind.Run, PowerPointParagraphInlineKind.LineBreak, PowerPointParagraphInlineKind.Run },
                inline.Select(node => node.Kind));
            Assert.Equal("one", inline[0].Text);
            Assert.Equal("two", inline[2].Text);
        }
    }

    [Theory]
    [InlineData("One\r\nTwo")]
    [InlineData("One\rTwo")]
    public void WindowsLineEndingsInLayoutTextBreakOnce(string text) {
        using var document = PowerPointDocument.Create();
        var options = PowerPointCompositionOptions.FromDesign(PowerPointDeckDesign.FromBrand("#2563EB", "text-breaks"));
        options.ApplyTheme = false;
        document.Compose(new PowerPointDeckPlan().AddContent("Text", new[] { new PowerPointLayoutText(text) }), options);
        using var saved = PowerPointDocument.Load(new MemoryStream(document.ToBytes()));
        var inline = saved.Slides[0].TextBoxes.SelectMany(box => box.Paragraphs).Single(paragraph => paragraph.Text.StartsWith("One")).InlineNodes;
        Assert.Equal(new[] { PowerPointParagraphInlineKind.Run, PowerPointParagraphInlineKind.LineBreak, PowerPointParagraphInlineKind.Run },
            inline.Select(node => node.Kind));
        Assert.Equal("One", inline[0].Text);
        Assert.Equal("Two", inline[2].Text);
    }

    [Fact]
    public void WrappedRowsOnASlideWithRoomForOneRowStillPaginate() {
        using var document = PowerPointDocument.Create();
        document.SlideSize.Width = OfficeLength.Inches(5);
        document.SlideSize.Height = OfficeLength.Inches(2);
        string wrapped = string.Join(" ", Enumerable.Repeat("wrapped cell text", 20));
        var table = new PowerPointLayoutTable(new[] { "Service", "Detail" },
            Enumerable.Range(0, 4).Select(_ => (IReadOnlyList<string>)new[] { wrapped, wrapped }).ToArray());
        var options = PowerPointCompositionOptions.FromDesign(PowerPointDeckDesign.FromBrand("#2563EB", "table-tiny"));
        options.ApplyTheme = false;
        var compose = System.Threading.Tasks.Task.Run(() => document.Compose(new PowerPointDeckPlan().AddContent("Services", new[] { table }), options));
        Assert.True(compose.Wait(TimeSpan.FromSeconds(30)), "Table pagination did not finish.");
        Assert.Equal(4, document.Slides.SelectMany(slide => slide.Tables).Sum(native => native.Rows - 1));
    }

    [Fact]
    public void RowsTallerThanTwoLinesPaginateInsideTheLiveArea() {
        using var document = PowerPointDocument.Create();
        var table = new PowerPointLayoutTable(new[] { "Service\vOwner", "Detail" },
            Enumerable.Range(0, 12).Select(index => (IReadOnlyList<string>)new[] { "Service " + index, "one\ntwo\nthree\nfour\nfive" }).ToArray());
        var options = PowerPointCompositionOptions.FromDesign(PowerPointDeckDesign.FromBrand("#2563EB", "table-tall"));
        options.ApplyTheme = false;
        document.Compose(new PowerPointDeckPlan().AddContent("Services", new[] { table }), options);
        var natives = document.Slides.SelectMany(slide => slide.Tables).ToArray();
        Assert.True(natives.Length > 1);
        Assert.Equal(12, natives.Sum(native => native.Rows - 1));
        foreach (var native in natives) {
            double size = native.GetCell(1, 1).Paragraphs[0].Runs[0].FontSize!.Value.ToPoints();
            Assert.True(native.Height.ToPoints() >= (native.Rows - 1) * 5 * size, "The frame must carry all five lines of each row.");
            Assert.True(native.Top.ToPoints() + native.Height.ToPoints() <= 488D, "The table must stay inside the live area.");
        }
    }

    [Fact]
    public void SavedTableUsesItsAuthoredCaptionAsAlternativeText() {
        const string caption = "Service & owners <review>";
        using var document = PowerPointDocument.Create();
        var table = new PowerPointLayoutTable(new[] { "Service", "Count" },
            new IReadOnlyList<string>[] { new[] { "Directory", "2" } }) { Caption = caption };
        var options = PowerPointCompositionOptions.FromDesign(PowerPointDeckDesign.FromBrand("#2563EB", "table-caption"));
        options.ApplyTheme = false;
        document.Compose(new PowerPointDeckPlan().AddContent("Services", new[] { table }), options);
        using var saved = PowerPointDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Equal(caption, Assert.Single(Assert.Single(saved.Slides).Tables).AltText);
        Assert.Empty(saved.ValidateDocument());
        Assert.DoesNotContain(saved.InspectAccessibility().Findings,
            finding => finding.Code == "accessibility-missing-alternative-text");
    }
}
