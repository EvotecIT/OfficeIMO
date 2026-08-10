using System;
using System.Linq;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfComposeIncrementalAdapterTests {
    [Fact]
    public void Defaults_ConfigureTopLevelFlowWithoutAddingAPageBoundary() {
        PdfDocument document = PdfDocument.Create(_ => { });

        document.Compose(compose => compose.Defaults(defaults => defaults
            .Size(300, 400)
            .Margin(24)
            .Background(PdfColor.FromRgb(240, 248, 255))));
        document.Compose(compose => compose.Content(content => content
            .Paragraph(paragraph => paragraph.Text("IncrementalAdapterMarker"))));

        byte[] bytes = document.ToBytes();
        string source = Encoding.ASCII.GetString(bytes);

        Assert.Contains("/MediaBox [0 0 300 400]", source, StringComparison.Ordinal);
        Assert.Contains("0.941 0.973 1 rg\n0 0 300 400 re f", source, StringComparison.Ordinal);
        Assert.Contains("IncrementalAdapterMarker", PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void Settings_UpdateTheDocumentOwnedOptionsSnapshot() {
        PdfDocument document = PdfDocument.Create(_ => { });

        document.Compose(compose => compose.Settings(settings =>
            settings.AddEmbeddedFile("evidence.txt", Encoding.UTF8.GetBytes("adapter evidence"), "text/plain")));
        document.Compose(compose => compose.Content(content => content
            .Paragraph(paragraph => paragraph.Text("Attachment report"))));

        byte[] bytes = document.ToBytes();
        var attachments = PdfDocument.Open(bytes).Read.Attachments();

        Assert.Single(attachments);
        Assert.Equal("evidence.txt", attachments.Single().FileName);
        Assert.Equal("adapter evidence", Encoding.UTF8.GetString(attachments.Single().Bytes));
    }

    [Fact]
    public void Settings_UpdateExistingExplicitPageSnapshots() {
        PdfDocument document = PdfDocument.Create(_ => { });
        document.Compose(compose => compose.Page(page => page
            .Size(300, 400)
            .Margin(24)
            .Content(content => content
                .Item(item => item.Paragraph(paragraph => paragraph.Text("LateSettingsMarker"))))));

        int callbackCount = 0;
        document.Compose(compose => compose.Settings(settings => {
            callbackCount++;
            settings.CompressContentStreams = false;
            settings.IncludeStandardFontToUnicodeMaps = true;
        }));

        string source = Encoding.ASCII.GetString(document.ToBytes());
        Assert.Equal(1, callbackCount);
        Assert.Contains("/MediaBox [0 0 300 400]", source, StringComparison.Ordinal);
        Assert.DoesNotContain("/Filter /FlateDecode", source, StringComparison.Ordinal);
        Assert.Contains("/ToUnicode", source, StringComparison.Ordinal);
        Assert.Contains("LateSettingsMarker", PdfReadDocument.Open(document.ToBytes()).ExtractText(), StringComparison.Ordinal);
    }
}
