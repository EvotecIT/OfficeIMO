using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Syntax;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfDestinationRegistryTests {
    [Fact]
    public void Registry_Classifies_Core_Destinations_And_Edit_Skip_Rules() {
        Assert.Equal(RtfDestinationType.Picture, RtfDestinationRegistry.GetDestinationType("pict"));
        Assert.Equal(RtfDestinationType.Object, RtfDestinationRegistry.GetDestinationType("object"));
        Assert.Equal(RtfDestinationType.Field, RtfDestinationRegistry.GetDestinationType("fldinst"));
        Assert.Equal(RtfDestinationType.Field, RtfDestinationRegistry.GetDestinationType("ffdata"));
        Assert.Equal(RtfDestinationType.Field, RtfDestinationRegistry.GetDestinationType("ffname"));
        Assert.Equal(RtfDestinationType.BodyText, RtfDestinationRegistry.GetDestinationType("upr"));
        Assert.Equal(RtfDestinationType.BodyText, RtfDestinationRegistry.GetDestinationType("ud"));
        Assert.Equal(RtfDestinationType.Footnote, RtfDestinationRegistry.GetDestinationType("footnote"));
        Assert.Equal(RtfDestinationType.Endnote, RtfDestinationRegistry.GetDestinationType("endnote"));
        Assert.Equal(RtfDestinationType.Annotation, RtfDestinationRegistry.GetDestinationType("annotation"));
        Assert.Equal(RtfDestinationType.Metadata, RtfDestinationRegistry.GetDestinationType("userprops"));
        Assert.Equal(RtfDestinationType.Metadata, RtfDestinationRegistry.GetDestinationType("docvar"));
        Assert.Equal(RtfDestinationType.Metadata, RtfDestinationRegistry.GetDestinationType("revtbl"));
        Assert.Equal(RtfDestinationType.Metadata, RtfDestinationRegistry.GetDestinationType("rsidtbl"));
        Assert.Equal(RtfDestinationType.Metadata, RtfDestinationRegistry.GetDestinationType("atnauthor"));
        Assert.Equal(RtfDestinationType.Metadata, RtfDestinationRegistry.GetDestinationType("filetbl"));
        Assert.Equal(RtfDestinationType.Metadata, RtfDestinationRegistry.GetDestinationType("file"));
        Assert.Equal(RtfDestinationType.Metadata, RtfDestinationRegistry.GetDestinationType("xmlnstbl"));
        Assert.Equal(RtfDestinationType.Metadata, RtfDestinationRegistry.GetDestinationType("xmlns"));
        Assert.Equal(RtfDestinationType.ListTable, RtfDestinationRegistry.GetDestinationType("listtext"));
        Assert.Equal(RtfDestinationType.Drawing, RtfDestinationRegistry.GetDestinationType("shp"));
        Assert.Equal(RtfDestinationType.Drawing, RtfDestinationRegistry.GetDestinationType("shptxt"));
        Assert.Equal(RtfDestinationType.Unknown, RtfDestinationRegistry.GetDestinationType("definitelyunknown"));

        Assert.True(RtfDestinationRegistry.ShouldSkipSemanticBinding("fonttbl"));
        Assert.True(RtfDestinationRegistry.ShouldSkipSemanticBinding("userprops"));
        Assert.True(RtfDestinationRegistry.ShouldSkipSemanticBinding("docvar"));
        Assert.True(RtfDestinationRegistry.ShouldSkipSemanticBinding("revtbl"));
        Assert.True(RtfDestinationRegistry.ShouldSkipSemanticBinding("rsidtbl"));
        Assert.True(RtfDestinationRegistry.ShouldSkipSemanticBinding("atnauthor"));
        Assert.True(RtfDestinationRegistry.ShouldSkipSemanticBinding("filetbl"));
        Assert.True(RtfDestinationRegistry.ShouldSkipSemanticBinding("xmlnstbl"));
        Assert.True(RtfDestinationRegistry.ShouldSkipSemanticBinding("listtext"));
        Assert.True(RtfDestinationRegistry.ShouldSkipSemanticBinding("shpinst"));
        Assert.False(RtfDestinationRegistry.ShouldSkipSemanticBinding("shp"));
        Assert.False(RtfDestinationRegistry.ShouldSkipSemanticBinding("shptxt"));
        Assert.False(RtfDestinationRegistry.ShouldSkipSemanticBinding("header"));
        Assert.False(RtfDestinationRegistry.ShouldSkipSemanticBinding("footnote"));
        Assert.False(RtfDestinationRegistry.ShouldSkipSemanticBinding("endnote"));
        Assert.False(RtfDestinationRegistry.IsUnsupportedSemanticDestination("header"));
        Assert.True(RtfDestinationRegistry.ShouldSkipTextReplacement("fldinst"));
        Assert.True(RtfDestinationRegistry.ShouldSkipSemanticBinding("ffdata"));
        Assert.True(RtfDestinationRegistry.ShouldSkipTextReplacement("ffdata"));
        Assert.True(RtfDestinationRegistry.ShouldSkipTextReplacement("ffname"));
        Assert.True(RtfDestinationRegistry.ShouldSkipTextReplacement("userprops"));
        Assert.True(RtfDestinationRegistry.ShouldSkipTextReplacement("docvar"));
        Assert.True(RtfDestinationRegistry.ShouldSkipTextReplacement("revtbl"));
        Assert.True(RtfDestinationRegistry.ShouldSkipTextReplacement("rsidtbl"));
        Assert.True(RtfDestinationRegistry.ShouldSkipTextReplacement("atnauthor"));
        Assert.True(RtfDestinationRegistry.ShouldSkipTextReplacement("filetbl"));
        Assert.True(RtfDestinationRegistry.ShouldSkipTextReplacement("xmlnstbl"));
        Assert.True(RtfDestinationRegistry.ShouldSkipTextReplacement("listtext"));
        Assert.True(RtfDestinationRegistry.ShouldSkipTextReplacement("shpinst"));
        Assert.False(RtfDestinationRegistry.ShouldSkipTextReplacement("shptxt"));
        Assert.False(RtfDestinationRegistry.ShouldSkipTextReplacement("fldrslt"));
        Assert.False(RtfDestinationRegistry.ShouldSkipSemanticBinding("object"));
        Assert.False(RtfDestinationRegistry.IsUnsupportedSemanticDestination("object"));
        Assert.True(RtfDestinationRegistry.IsUnsupportedSemanticDestination("objdata"));
        Assert.True(RtfDestinationRegistry.ShouldSkipTextReplacement("objdata"));
    }

    [Fact]
    public void Registry_Detects_Ignorable_Destination_Groups() {
        RtfSyntaxTree tree = RtfSyntaxTree.Parse(@"{\rtf1\ansi{\*\unknown Hidden}\pard Visible\par}");
        RtfGroup unknownGroup = tree.Root.Children.OfType<RtfGroup>().Single(group => group.Destination == "unknown");

        Assert.True(RtfDestinationRegistry.IsIgnorableDestinationGroup(unknownGroup));
    }

    [Theory]
    [InlineData("themedata", RtfDestinationType.Metadata)]
    [InlineData("colorschememapping", RtfDestinationType.Metadata)]
    [InlineData("latentstyles", RtfDestinationType.StyleSheet)]
    [InlineData("xe", RtfDestinationType.Field)]
    [InlineData("tc", RtfDestinationType.Field)]
    [InlineData("xmlopen", RtfDestinationType.Metadata)]
    [InlineData("factoidname", RtfDestinationType.Metadata)]
    [InlineData("datastore", RtfDestinationType.Metadata)]
    [InlineData("protstart", RtfDestinationType.Metadata)]
    [InlineData("mvfmf", RtfDestinationType.Bookmark)]
    public void Registry_Classifies_Preserve_Only_Advanced_Families(string destination, RtfDestinationType type) {
        Assert.Equal(type, RtfDestinationRegistry.GetDestinationType(destination));
        Assert.True(RtfDestinationRegistry.ShouldSkipSemanticBinding(destination));
        Assert.True(RtfDestinationRegistry.ShouldSkipTextReplacement(destination));
        Assert.True(RtfDestinationRegistry.IsUnsupportedSemanticDestination(destination));
    }

    [Fact]
    public void Semantic_Read_Diagnoses_Known_Preserve_Only_Families_And_Keeps_Visible_Text() {
        const string rtf = @"{\rtf1\ansi{\*\themedata 001122}{\*\xe Index entry}{\*\xmlopen\xmlns1}{\*\xmlclose}\pard Visible\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal("Visible", Assert.Single(result.Document.Paragraphs).ToPlainText());
        Assert.Equal(rtf, result.ToRtfLossless());
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "RTF101" && diagnostic.Message.Contains("themedata"));
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "RTF101" && diagnostic.Message.Contains("xe"));
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "RTF101" && diagnostic.Message.Contains("xmlopen"));
    }
}
