using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfInspectorTests {
    private static void AssertNamedDestination(PdfDocumentInfo info, string name, int pageNumber, double destinationTop) {
        Assert.Equal(1, info.NamedDestinationCount);
        Assert.Equal(new[] { name }, info.NamedDestinationNames);

        PdfNamedDestination destination = Assert.Single(info.NamedDestinations);
        Assert.Equal(name, destination.Name);
        Assert.Equal(pageNumber, destination.PageNumber);
        Assert.Equal(destinationTop, destination.DestinationTop);
    }

    private static void AssertOpenAction(PdfDocumentInfo info, string actionType, int pageNumber, double? destinationTop) {
        Assert.True(info.HasReadableOpenAction);
        Assert.NotNull(info.OpenAction);
        Assert.Equal(actionType, info.OpenAction!.ActionType);
        Assert.Equal(pageNumber, info.OpenAction.PageNumber);
        Assert.Equal(destinationTop, info.OpenAction.DestinationTop);
    }

    private static void AssertViewerPreferences(PdfDocumentInfo info) {
        Assert.True(info.HasReadableViewerPreferences);
        Assert.NotNull(info.ViewerPreferences);
        Assert.Equal(2, info.ViewerPreferences!.Count);
        Assert.Equal("true", info.ViewerPreferences.GetValue("HideToolbar"));
        Assert.Equal("true", info.ViewerPreferences.GetValue("DisplayDocTitle"));
        Assert.True(info.ViewerPreferences.GetBoolean("HideToolbar"));
        Assert.True(info.ViewerPreferences.GetBoolean("DisplayDocTitle"));
        Assert.Null(info.ViewerPreferences.GetValue("Missing"));
        Assert.Null(info.ViewerPreferences.GetBoolean("Missing"));
    }

    private static void AssertPageLabel(PdfDocumentInfo info, int startPageIndex, int startPageNumber, string? style, string? prefix, int? startNumber) {
        Assert.True(info.HasReadablePageLabels);
        Assert.Equal(1, info.PageLabelCount);

        PdfPageLabel label = Assert.Single(info.PageLabels);
        Assert.Equal(startPageIndex, label.StartPageIndex);
        Assert.Equal(startPageNumber, label.StartPageNumber);
        Assert.Equal(style, label.Style);
        Assert.Equal(prefix, label.Prefix);
        Assert.Equal(startNumber, label.StartNumber);
    }

    private static void AssertRewriteBlocker(PdfDocumentPreflight report, PdfRewriteBlockerKind kind, string message) {
        PdfRewriteBlocker? blocker = null;
        for (int i = 0; i < report.RewriteBlockers.Count; i++) {
            if (report.RewriteBlockers[i].Kind == kind) {
                blocker = report.RewriteBlockers[i];
                break;
            }
        }

        Assert.NotNull(blocker);
        Assert.Equal(message, blocker!.Message);
        Assert.True(report.HasRewriteBlocker(kind));
    }

    private static void AssertReadBlocker(PdfDocumentPreflight report, PdfReadBlockerKind kind, string message) {
        PdfReadBlocker? blocker = null;
        for (int i = 0; i < report.ReadBlockers.Count; i++) {
            if (report.ReadBlockers[i].Kind == kind) {
                blocker = report.ReadBlockers[i];
                break;
            }
        }

        Assert.NotNull(blocker);
        Assert.Equal(message, blocker!.Message);
        Assert.True(report.HasReadBlocker(kind));
    }


}
