using OfficeIMO.Html;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests.Rtf;

public class RtfColorSlotRegressionTests {
    [Fact]
    public void Explicit_First_Color_Is_Remapped_Consistently_Without_Changing_Lossless_Source() {
        const string input = @"{\rtf1\ansi{\colortbl\red255\green0\blue0;\red0\green0\blue255;}\cf0 Red\cf1 Blue\par}";
        RtfReadResult read = RtfDocument.Read(input);
        RtfDocument document = read.Document;
        Assert.Equal(input, read.ToRtfLossless());
        Assert.Equal("#FF0000", document.GetColor(document.Paragraphs[0].Runs[0].ForegroundColorIndex!.Value)!.ToString());
        Assert.Equal("#0000FF", document.GetColor(document.Paragraphs[0].Runs[1].ForegroundColorIndex!.Value)!.ToString());
        string html = document.ToHtml();
        Assert.Contains("#FF0000", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("#0000FF", html, StringComparison.OrdinalIgnoreCase);
        using WordDocument word = document.ToWordDocument();
        Assert.Contains(word.Paragraphs, run => run.Text == "Blue" && run.ColorHex == "0000FF");
        PdfCore.RichParagraphBlock pdfParagraph = Assert.IsType<PdfCore.RichParagraphBlock>(Assert.Single(document.ToPdfDocument().Blocks));
        Assert.Equal(PdfCore.PdfColor.FromRgb(0, 0, 255), pdfParagraph.Runs.Last().Color);
        RtfDocument reopened = RtfDocument.Read(document.ToRtf()).Document;
        Assert.Equal("#0000FF", reopened.GetColor(reopened.Paragraphs[0].Runs[1].ForegroundColorIndex!.Value)!.ToString());
    }

    [Fact]
    public void Consecutive_Empty_Slots_Retain_Color_Identity_In_Save_Html_Metadata_And_Append() {
        RtfDocument source = RtfDocument.Read(@"{\rtf1\ansi{\colortbl;\red255\green0\blue0;;;\red0\green0\blue255;}\cf4 Blue\cf2 Auto\par}").Document;
        Assert.Equal(4, source.Colors.Count);
        Assert.True(source.Colors[1].IsAutomatic);
        Assert.True(source.Colors[2].IsAutomatic);
        Assert.Null(source.GetColor(2));
        Assert.Equal("#0000FF", source.GetColor(4)!.ToString());
        RtfDocument saved = RtfDocument.Read(source.ToRtf()).Document;
        Assert.Equal("#0000FF", saved.GetColor(4)!.ToString());
        Assert.True(saved.Colors[2].IsAutomatic);
        RtfDocument htmlRoundTrip = HtmlConversionDocument.Parse(source.ToHtml(new RtfToHtmlOptions { IncludeRoundTripMetadata = true, FragmentOnly = false })).ToRtfDocument();
        Assert.True(htmlRoundTrip.Colors[1].IsAutomatic);
        Assert.Equal("#0000FF", htmlRoundTrip.GetColor(4)!.ToString());

        RtfDocument destination = RtfDocument.Create();
        destination.AddColor(0, 255, 0);
        destination.AppendDocument(source);
        RtfRun blue = destination.Paragraphs[0].Runs[0];
        Assert.Equal("#0000FF", destination.GetColor(blue.ForegroundColorIndex!.Value)!.ToString());
        Assert.Null(destination.GetColor(destination.Paragraphs[0].Runs[1].ForegroundColorIndex!.Value));
    }
}
