using OfficeIMO.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public sealed class RtfMergeDocumentSettingsRegressionTests {
    [Fact]
    public void AHeaderOnlySectionImportInvalidatesExistingPageCounts() {
        RtfDocument destination = RtfDocument.Create();
        destination.AddParagraph("Destination");
        destination.Info.NumberOfPages = 1;
        RtfDocument source = RtfDocument.Create();
        source.AddHeader().AddParagraph("Imported header");

        var result = destination.AppendDocument(source, new RtfDocumentMergeOptions { PreserveSections = true });
        Assert.Equal(2, destination.Sections.Count);
        Assert.Null(destination.Info.NumberOfPages);
        Assert.Contains(result.Report.Diagnostics, item => item.Code == "RtfMergeStatisticsInvalidated");
        Assert.Throws<RtfConversionLossException>(() => result.Report.RequireNoLoss());
    }

    [Fact]
    public void AppendReportsInvalidationOfCurrentDestinationAlternateHtml() {
        RtfDocument destination = RtfDocument.Read(@"{\rtf1\ansi\fromhtml1{\*\htmltag <p>Original</p>}\htmlrtf1 Original\par}").Document;
        Assert.True(destination.IsHtmlEncapsulationCurrent);
        RtfDocument source = RtfDocument.Create();
        source.AddParagraph("Appended");

        var result = destination.AppendDocument(source);
        Assert.False(destination.IsHtmlEncapsulationCurrent);
        Assert.Contains(result.Report.Diagnostics, item => item.Code == "RtfMergeHtmlEncapsulationOmitted");
        Assert.Throws<RtfConversionLossException>(() => result.Report.RequireNoLoss());
    }

    [Fact]
    public void AnEmptyDestinationAdoptsSourceGlobalSettingsButACompositeReportsConflicts() {
        RtfDocument source = RtfDocument.Create();
        source.Settings.DefaultTabWidthTwips = 360;
        source.Settings.FacingPages = true;
        source.Settings.ReadOnlyProtection = true;
        source.AddParagraph("Imported\tcontent");
        RtfDocument empty = RtfDocument.Create();
        empty.AppendDocument(source).Report.RequireNoLoss();
        Assert.Equal(360, empty.Settings.DefaultTabWidthTwips);
        Assert.True(empty.Settings.FacingPages);
        Assert.True(empty.Settings.ReadOnlyProtection);

        RtfDocument destination = RtfDocument.Create();
        destination.Settings.DefaultTabWidthTwips = 720;
        destination.Settings.FacingPages = false;
        destination.Settings.ReadOnlyProtection = false;
        destination.AddParagraph("Destination");
        var result = destination.AppendDocument(source, new RtfDocumentMergeOptions { PreserveSections = true });
        var loss = Assert.Single(result.Report.Diagnostics, item => item.Code == "RtfMergeDocumentSettingsFlattened");
        Assert.Equal(3, loss.Count);
        Assert.Throws<RtfConversionLossException>(() => result.Report.RequireNoLoss());
        Assert.Equal(720, destination.Settings.DefaultTabWidthTwips);
        Assert.False(destination.Settings.FacingPages);
        Assert.False(destination.Settings.ReadOnlyProtection);
        Assert.True(source.Settings.ReadOnlyProtection);
    }

    [Fact]
    public void AppendInvalidatesAggregateCountsOnlyWhenCombiningContent() {
        RtfDocument source = RtfDocument.Create();
        source.AddParagraph("Two words");
        source.Info.NumberOfWords = 2;
        RtfDocument destination = RtfDocument.Create();
        destination.AddParagraph("One");
        destination.Info.NumberOfPages = 1;
        destination.Info.NumberOfWords = 1;
        destination.Info.NumberOfCharacters = 3;
        destination.Info.NumberOfCharactersWithSpaces = 3;
        var result = destination.AppendDocument(source);
        Assert.Contains(result.Report.Diagnostics, item => item.Code == "RtfMergeStatisticsInvalidated" && item.Action == RtfConversionAction.Omitted);
        Assert.Null(destination.Info.NumberOfPages);
        Assert.Null(destination.Info.NumberOfWords);
        Assert.Null(destination.Info.NumberOfCharacters);
        Assert.Null(destination.Info.NumberOfCharactersWithSpaces);
        Assert.Equal(2, source.Info.NumberOfWords);
        RtfDocument empty = RtfDocument.Create();
        empty.AppendDocument(source).Report.RequireNoLoss();
        Assert.Equal(2, empty.Info.NumberOfWords);
    }

    [Fact]
    public void FlatteningRootPageSetupAndAlternateHtmlIsExplicitInTheMergeReport() {
        RtfDocument source = RtfDocument.Read(@"{\rtf1\ansi\fromhtml1{\*\htmltag <p>Alternate</p>}\htmlrtf1 Body\par}").Document;
        source.PageSetup.MarginLeftTwips = 2880;
        var result = RtfDocument.Create().AppendDocument(source);
        Assert.Contains(result.Report.Diagnostics, item => item.Code == "RtfMergeDocumentPageSetupOmitted" && item.Action == RtfConversionAction.Omitted);
        Assert.Contains(result.Report.Diagnostics, item => item.Code == "RtfMergeHtmlEncapsulationOmitted" && item.Action == RtfConversionAction.Omitted);
        Assert.Throws<RtfConversionLossException>(() => result.Report.RequireNoLoss());
    }
}
