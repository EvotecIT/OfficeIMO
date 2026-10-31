using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageFootnoteTests {
    [Theory]
    [InlineData(PdfFlowOverflowBehavior.Skip, true)]
    [InlineData(PdfFlowOverflowBehavior.StopDocument, false)]
    public void FlowContinuationOverflowUsesBodyCapacityAfterMinimumHeightAdvance(
        PdfFlowOverflowBehavior behavior, bool followingBodyVisible) {
        PdfDocument document = CreateContinuationCapacityDocument();
        var capture = new PdfLayoutPositionCapture();
        document.Content.Flow(body => {
            for (int index = 0; index < 6; index++) body.Text("BoundedBody" + index, style: Style());
        }, new PdfFlowOptions { MinimumRemainingHeight = global::OfficeIMO.OfficeLength.Points(10), OverflowBehavior = behavior }, capture);
        document.Content.Text("AfterBoundedFlow", style: Style());

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        string text = string.Concat(pdf.GetPages().Select(page => page.Text));
        Assert.True(capture.WasSkipped);
        Assert.Empty(capture.Regions);
        for (int index = 0; index < 6; index++) Assert.DoesNotContain("BoundedBody" + index, text, StringComparison.Ordinal);
        Assert.Equal(followingBodyVisible, text.Contains("AfterBoundedFlow", StringComparison.Ordinal));
        AssertContinuationCapacityNoteInventory(pdf);
        if (followingBodyVisible) Assert.Contains("AfterBoundedFlow", pdf.GetPage(3).Text, StringComparison.Ordinal);
    }

    [Fact]
    public void FlowContinuationMinimumHeightAdvancesPastTheLastPartiallyReservedPage() {
        PdfDocument document = CreateContinuationCapacityDocument();
        var capture = new PdfLayoutPositionCapture();
        // The last continued-note page has 40pt of body room. A 50pt minimum
        // therefore requires the next unreserved page even for one 10pt line.
        document.Content.Flow(body => body.Text("MinimumHeightBody", style: Style()),
            new PdfFlowOptions { MinimumRemainingHeight = global::OfficeIMO.OfficeLength.Points(50) }, capture);

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        AssertContinuationCapacityNoteInventory(pdf);
        Assert.Equal(4, pdf.NumberOfPages);
        Assert.DoesNotContain("MinimumHeightBody", pdf.GetPage(3).Text, StringComparison.Ordinal);
        Assert.Contains("MinimumHeightBody", pdf.GetPage(4).Text, StringComparison.Ordinal);
        Assert.Equal(4, Assert.Single(capture.Regions).PageNumber);
        Assert.False(capture.WasSkipped);
    }

    [Fact]
    public void FlowContinuationCallbacksReportRemainingBodyRoomAndTheFinalPhysicalPage() {
        PdfDocument document = CreateContinuationCapacityDocument();
        var predicates = new List<PdfFlowContext>();
        var factories = new List<PdfFlowContext>();
        var capture = new PdfLayoutPositionCapture();
        document.Content.Deferred(context => {
            factories.Add(context);
            return body => body.Text("ContextPage" + context.PageNumber + "AtTop" + context.IsAtPageTop, style: Style());
        }, new PdfFlowOptions {
            MinimumRemainingHeight = global::OfficeIMO.OfficeLength.Points(20),
            ShowIf = context => { predicates.Add(context); return true; }
        }, capture);
        var followingContexts = new List<PdfFlowContext>();
        document.Content.Flow(body => body.Text("AfterContextBody", style: Style()), new PdfFlowOptions {
            ShowIf = context => { followingContexts.Add(context); return !context.IsAtPageTop; }
        });

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        AssertContinuationCapacityNoteInventory(pdf);
        Assert.Contains("ContextPage3AtTopTrue", pdf.GetPage(3).Text, StringComparison.Ordinal);
        Assert.Contains("AfterContextBody", pdf.GetPage(3).Text, StringComparison.Ordinal);
        Assert.DoesNotContain("ContextPage2", string.Concat(pdf.GetPages().Select(page => page.Text)), StringComparison.Ordinal);
        Assert.Contains(predicates, context => context.PageNumber == 1 && Math.Abs(context.AvailableHeight.ToPoints()) < .001D);
        Assert.Contains(factories, context => context.PageNumber == 1 && Math.Abs(context.AvailableHeight.ToPoints()) < .001D);
        Assert.Contains(factories, context => context.PageNumber == 2 && Math.Abs(context.AvailableHeight.ToPoints()) < .001D);
        Assert.Contains(factories, context => context.PageNumber == 3 && Math.Abs(context.AvailableHeight.ToPoints() - 40D) < .001D);
        Assert.All(factories, context => Assert.Equal(100D, context.FullContentHeight.ToPoints(), 3));
        Assert.All(predicates, context => Assert.False(context.IsAtPageTop));
        Assert.All(factories, context => Assert.Equal(context.PageNumber > 1, context.IsAtPageTop));
        Assert.All(followingContexts, context => {
            Assert.Equal(3, context.PageNumber);
            Assert.Equal(30D, context.AvailableHeight.ToPoints(), 3);
            Assert.False(context.IsAtPageTop);
        });
        Assert.NotEmpty(followingContexts);
        Assert.Equal(3, Assert.Single(capture.Regions).PageNumber);
    }

    private static PdfDocument CreateContinuationCapacityDocument() {
        PdfDocument document = PdfDocument.Create(Options());
        PdfFootnoteDefinition note = document.Content.DefineFootnote(body => {
            for (int index = 0; index < 25; index++) body.Text("CapacityNote" + index.ToString("00"), style: Style());
        });
        document.Content.Paragraph(paragraph => paragraph.Runs(new[] {
            PdfTextRun.Normal("CapacityReference").WithFootnoteReference(note)
        }), style: Style());
        return document;
    }

    private static void AssertContinuationCapacityNoteInventory(PdfPigDocument pdf) {
        string text = string.Concat(pdf.GetPages().Select(page => page.Text));
        for (int index = 0; index < 25; index++) {
            string marker = "CapacityNote" + index.ToString("00");
            Assert.Single(Regex.Matches(text, Regex.Escape(marker)));
            int expectedPage = index < 9 ? 1 : index < 19 ? 2 : 3;
            Assert.Contains(marker, pdf.GetPage(expectedPage).Text, StringComparison.Ordinal);
        }
        Assert.Single(Regex.Matches(text, "CapacityReference"));
    }
}
