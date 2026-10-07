using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(0)]
    [InlineData(30)]
    [InlineData(100)]
    public void HtmlColumnNotes_ReserveOnlyTheirOriginatingColumnWithoutBodyOverlap(int precedingHeight) {
        string html = ColumnNoteStyle(160)
            + "<div style='height:" + precedingHeight + "px'></div><section class='columns'>"
            + "<p>CallA<span class='note'>NoteA</span></p>"
            + string.Concat(Enumerable.Range(0, 8).Select(i => "<p>Body" + i.ToString("D2") + "</p>"))
            + "<p>CallB<span class='note'>NoteB</span></p>"
            + string.Concat(Enumerable.Range(8, 4).Select(i => "<p>Body" + i.ToString("D2") + "</p>"))
            + "</section><p>End</p>";
        HtmlRenderDocument document = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged, HonorCssPageRules = true
        });
        for (int n = 1; n <= 2; n++) {
            string noteText = n == 1 ? "NoteA" : "NoteB";
            HtmlRenderPage callPage = Assert.Single(document.Pages, p => p.Visuals.OfType<HtmlRenderNamedDestination>()
                .Any(d => d.Name == "officeimo-footnote-call-" + n));
            HtmlRenderNamedDestination call = Assert.Single(callPage.Visuals.OfType<HtmlRenderNamedDestination>(),
                d => d.Name == "officeimo-footnote-call-" + n);
            HtmlRenderText note = Assert.Single(callPage.Visuals.OfType<HtmlRenderText>(), t => t.Text == noteText);
            double left = call.X < 130D ? 10D : 130D;
            Assert.InRange(note.X, left, left + 100D);
            Assert.True(note.X + note.Width <= left + 100D + 0.001D);
            foreach (HtmlRenderText body in callPage.Visuals.OfType<HtmlRenderText>()
                .Where(t => (t.Text.StartsWith("Body") || t.Text.StartsWith("Call")) && t.X >= left && t.X < left + 100D)) {
                Assert.True(body.Y + body.Height <= note.Y + 0.001D);
            }
        }
        for (int i = 0; i < 12; i++) {
            Assert.Single(document.Pages.SelectMany(p => p.Visuals).OfType<HtmlRenderText>(), t => t.Text == "Body" + i.ToString("D2"));
        }
        Assert.Single(document.Pages.SelectMany(p => p.Visuals).OfType<HtmlRenderText>(), t => t.Text == "End");
        Assert.DoesNotContain(document.Diagnostics, d => d.Severity == HtmlDiagnosticSeverity.Error);
    }

    [Fact]
    public void HtmlColumnNotes_LongNoteContinuesWithAllLinesAndNavigation() {
        string html = ColumnNoteStyle(100)
            + "<section class='columns'><p>Call<span class='note'>"
            + string.Concat(Enumerable.Range(0, 24).Select(i => "Long" + i.ToString("D2") + "<br>"))
            + "</span></p>"
            + string.Concat(Enumerable.Range(0, 16).Select(i => "<p>Body" + i.ToString("D2") + "</p>"))
            + "</section><p>End</p>";
        HtmlRenderDocument document = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged, HonorCssPageRules = true
        });
        HtmlRenderText[] text = document.Pages.SelectMany(p => p.Visuals).OfType<HtmlRenderText>().ToArray();
        for (int i = 0; i < 24; i++) Assert.Single(text, t => t.Text == "Long" + i.ToString("D2"));
        for (int i = 0; i < 16; i++) Assert.Single(text, t => t.Text == "Body" + i.ToString("D2"));
        Assert.Single(text, t => t.Text == "End");
        Assert.Contains(text, t => t.Text.Contains("cont."));
        HtmlRenderPage first = document.Pages.First();
        HtmlRenderText firstNote = Assert.Single(first.Visuals.OfType<HtmlRenderText>(), t => t.Text == "Long00");
        Assert.True(firstNote.X + firstNote.Width <= 110D + 0.001D);
        Assert.All(first.Visuals.OfType<HtmlRenderText>().Where(t => t.Text.StartsWith("Long")),
            t => Assert.True(t.Y + t.Height <= 110D + 0.001D));
        Assert.All(document.Pages, p => Assert.All(p.Visuals.OfType<HtmlRenderText>(), t => {
            Assert.True(t.X >= 10D - 0.001D && t.X + t.Width <= 230D + 0.001D);
            Assert.True(t.Y >= 10D - 0.001D && t.Y + t.Height <= 230D + 0.001D);
        }));
        HtmlRenderNamedDestination[] destinations = document.Pages.SelectMany(p => p.Visuals).OfType<HtmlRenderNamedDestination>().ToArray();
        Assert.Single(destinations, d => d.Name == "officeimo-footnote-call-1");
        Assert.Single(destinations, d => d.Name == "officeimo-footnote-note-1");
    }

    [Fact]
    public void HtmlColumnNotes_AtomicBodyImageDoesNotOverlapReservedNote() {
        const string pixel = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNgYAAAAAMAASsJTYQAAAAASUVORK5CYII=";
        string html = ColumnNoteStyle(100) + "<section class='columns'><p>Call<span class='note'>NoteBody</span></p>"
            + "<img style='display:block;width:50px;height:95px' src='data:image/png;base64," + pixel + "'>"
            + "<p>AfterImage</p></section><p>End</p>";
        HtmlRenderDocument document = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged, HonorCssPageRules = true,
            ResourceUrlPolicy = HtmlUrlPolicy.CreateEmbeddedResourceProfile()
        });
        HtmlRenderImage image = Assert.Single(document.Pages.SelectMany(p => p.Visuals).OfType<HtmlRenderImage>());
        Assert.Equal(95D, image.Height, 3);
        HtmlRenderPage imagePage = Assert.Single(document.Pages, p => p.Visuals.Contains(image));
        foreach (HtmlRenderText note in imagePage.Visuals.OfType<HtmlRenderText>().Where(t => t.Text == "NoteBody")) {
            Assert.True(image.X + image.Width <= note.X || note.X + note.Width <= image.X
                || image.Y + image.Height <= note.Y || note.Y + note.Height <= image.Y);
        }
        Assert.Single(document.Pages.SelectMany(p => p.Visuals).OfType<HtmlRenderText>(), t => t.Text == "AfterImage");
        Assert.DoesNotContain(document.Diagnostics, d => d.Severity == HtmlDiagnosticSeverity.Error);
    }

    [Fact]
    public void HtmlColumnNotes_AtomicBodyImageAvoidsContinuationReservations() {
        const string pixel = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNgYAAAAAMAASsJTYQAAAAASUVORK5CYII=";
        string html = ColumnNoteStyle(100) + "<section class='columns'><p>Call<span class='note'>"
            + string.Concat(Enumerable.Range(0, 12).Select(i => "Long" + i.ToString("D2") + "<br>"))
            + "</span></p><img style='display:block;width:50px;height:95px' src='data:image/png;base64," + pixel + "'>"
            + "<p>AfterImage</p></section><p>End</p>";
        HtmlRenderDocument document = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged, HonorCssPageRules = true,
            ResourceUrlPolicy = HtmlUrlPolicy.CreateEmbeddedResourceProfile()
        });
        HtmlRenderImage image = Assert.Single(document.Pages.SelectMany(p => p.Visuals).OfType<HtmlRenderImage>());
        Assert.Equal(95D, image.Height, 3);
        HtmlRenderPage page = Assert.Single(document.Pages, p => p.Visuals.Contains(image));
        Assert.All(page.Visuals.OfType<HtmlRenderText>().Where(t => t.Text.StartsWith("Long")), note => {
            Assert.True(image.X + image.Width <= note.X || note.X + note.Width <= image.X
                || image.Y + image.Height <= note.Y || note.Y + note.Height <= image.Y);
        });
        for (int i = 0; i < 12; i++) Assert.Single(document.Pages.SelectMany(p => p.Visuals).OfType<HtmlRenderText>(),
            t => t.Text == "Long" + i.ToString("D2"));
    }

    [Fact]
    public void HtmlColumnNotes_DeferredCallPreservesParagraphWidows() {
        string html = ColumnNoteStyle(100) + "<section class='columns'><p style='widows:2;orphans:2'>"
            + "Line00<br>Line01<br>Line02<br>Line03<br>Line04<span class='note'>Note00<br>Note01<br>Note02</span>"
            + "</p></section><p>End</p>";
        HtmlRenderDocument document = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged, HonorCssPageRules = true
        });
        HtmlRenderPage lastPage = Assert.Single(document.Pages, p => p.Visuals.OfType<HtmlRenderText>().Any(t => t.Text == "Line04"));
        HtmlRenderText last = Assert.Single(lastPage.Visuals.OfType<HtmlRenderText>(), t => t.Text == "Line04");
        HtmlRenderText previous = Assert.Single(lastPage.Visuals.OfType<HtmlRenderText>(), t => t.Text == "Line03");
        Assert.Equal(previous.X, last.X, 3);
        Assert.Equal(20D, last.Y - previous.Y, 3);
        Assert.Equal(5, document.Pages.SelectMany(p => p.Visuals).OfType<HtmlRenderText>().Count(t => t.Text.StartsWith("Line")));
    }

    [Fact]
    public void HtmlColumnNotes_AtomicNoteImageContinuesWholeWithoutLosingContent() {
        const string pixel = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNgYAAAAAMAASsJTYQAAAAASUVORK5CYII=";
        string html = ColumnNoteStyle(100) + "<section class='columns'><p>Call<span class='note'>NoteIntro<br>"
            + "<img style='display:block;width:50px;height:80px' src='data:image/png;base64," + pixel + "'>NoteEnd</span></p>"
            + "<p>Body</p></section><p>End</p>";
        HtmlRenderDocument document = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged, HonorCssPageRules = true,
            ResourceUrlPolicy = HtmlUrlPolicy.CreateEmbeddedResourceProfile()
        });
        HtmlRenderImage image = Assert.Single(document.Pages.SelectMany(p => p.Visuals).OfType<HtmlRenderImage>());
        Assert.Equal(80D, image.Height, 3);
        Assert.InRange(image.Y, 10D, 30D);
        HtmlRenderText[] text = document.Pages.SelectMany(p => p.Visuals).OfType<HtmlRenderText>().ToArray();
        foreach (string label in new[] { "NoteIntro", "NoteEnd", "Body", "End" }) Assert.Single(text, t => t.Text == label);
        Assert.DoesNotContain(document.Diagnostics, d => d.Code == HtmlRenderDiagnosticCodes.VisualFragmentUnsupported);
    }

    [Fact]
    public void HtmlColumnNotes_ContinuationHonorsTheConfiguredColumnLimit() {
        string html = ColumnNoteStyle(100) + "<section class='columns'><p>Call<span class='note'>"
            + string.Concat(Enumerable.Range(0, 24).Select(i => "Note" + i.ToString("D2") + "<br>"))
            + "</span></p></section>";
        HtmlDomLimitException error = Assert.Throws<HtmlDomLimitException>(() => HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged, HonorCssPageRules = true, MaxColumnCount = 2 }));
        Assert.Equal(HtmlRenderDiagnosticCodes.MultiColumnLimitExceeded, error.Code);
        Assert.Equal(nameof(HtmlRenderOptions.MaxColumnCount), error.LimitSource);
    }

    [Theory]
    [InlineData(40)]
    [InlineData(60)]
    public void HtmlColumnNotes_ReservationFitsTheMeasuredCallLine(int lineHeight) {
        string html = ColumnNoteStyle(100) + "<section class='columns'><p style='line-height:" + lineHeight + "px'>Call"
            + "<span class='note'>" + string.Concat(Enumerable.Range(0, 12).Select(i => "Note" + i.ToString("D2") + "<br>"))
            + "</span></p></section><p>End</p>";
        HtmlRenderDocument document = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged, HonorCssPageRules = true
        });
        HtmlRenderPage callPage = Assert.Single(document.Pages, p => p.Visuals.OfType<HtmlRenderNamedDestination>()
            .Any(d => d.Name == "officeimo-footnote-call-1"));
        HtmlRenderText call = Assert.Single(callPage.Visuals.OfType<HtmlRenderText>(), t => t.Text == "Call");
        HtmlRenderText note = Assert.Single(callPage.Visuals.OfType<HtmlRenderText>(), t => t.Text == "Note00");
        Assert.True(call.Y + lineHeight <= note.Y + 0.001D);
        Assert.InRange(note.X, call.X, call.X + 100D);
        for (int i = 0; i < 12; i++) Assert.Single(document.Pages.SelectMany(p => p.Visuals).OfType<HtmlRenderText>(),
            t => t.Text == "Note" + i.ToString("D2"));
        Assert.DoesNotContain(document.Diagnostics, d => d.Severity == HtmlDiagnosticSeverity.Error);
    }

    [Fact]
    public void HtmlPagedMedia_UnnamedFootnotesStayWithTheirDistinctCallPages() {
        string html = "<style>@page{size:240px 160px;margin:10px}body,p{margin:0;font:12px/16px Arial}"
            + ".note{float:footnote;font:10px/12px Arial}</style>"
            + "<p>First<span class='note'>NoteOne</span></p>"
            + "<p style='break-before:page'>Second<span class='note'>NoteTwo</span></p>";
        HtmlRenderDocument document = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged, HonorCssPageRules = true
        });
        for (int number = 1; number <= 2; number++) {
            HtmlRenderPage call = Assert.Single(document.Pages, p => p.Visuals.OfType<HtmlRenderNamedDestination>()
                .Any(d => d.Name == "officeimo-footnote-call-" + number));
            HtmlRenderPage note = Assert.Single(document.Pages, p => p.Visuals.OfType<HtmlRenderText>()
                .Any(t => t.Text == (number == 1 ? "NoteOne" : "NoteTwo")));
            Assert.Equal(call.PageNumber, note.PageNumber);
        }
    }

    private static string ColumnNoteStyle(int height) =>
        "<style>@page{size:240px 240px;margin:10px}body{margin:0;font:10px/20px Arial}p{margin:0}"
        + ".columns{column-count:2;column-gap:20px;column-fill:auto;height:" + height + "px}"
        + ".note{float:footnote;float-reference:column;font:10px/12px Arial}</style>";
}
