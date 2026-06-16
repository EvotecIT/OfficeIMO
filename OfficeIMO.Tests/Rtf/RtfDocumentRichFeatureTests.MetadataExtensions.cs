using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Diagnostics;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public partial class RtfDocumentRichFeatureTests {
    [Fact]
    public void Write_And_Read_User_Properties() {
        RtfDocument document = RtfDocument.Create();
        document.AddUserProperty(RtfUserProperty.Text("Client", "Contoso"));
        document.AddUserProperty(RtfUserProperty.Boolean("Approved", true));
        document.AddUserProperty(RtfUserProperty.Number("Score", 98.5));
        RtfUserProperty linked = document.AddUserProperty("External");
        linked.LinkedValue = "Sheet1!A1";
        document.AddParagraph("Body");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"{\*\userprops{\propname Client}\proptype30{\staticval Contoso}{\propname Approved}\proptype11{\staticval 1}{\propname Score}\proptype5{\staticval 98.5}{\propname External}{\linkval Sheet1!A1}}", rtf, StringComparison.Ordinal);
        Assert.Collection(read.Document.UserProperties,
            property => {
                Assert.Equal("Client", property.Name);
                Assert.Equal(RtfUserProperty.TextType, property.TypeCode);
                Assert.Equal("Contoso", property.StaticValue);
            },
            property => {
                Assert.Equal("Approved", property.Name);
                Assert.Equal(RtfUserProperty.BooleanType, property.TypeCode);
                Assert.Equal("1", property.StaticValue);
            },
            property => {
                Assert.Equal("Score", property.Name);
                Assert.Equal(RtfUserProperty.NumberType, property.TypeCode);
                Assert.Equal("98.5", property.StaticValue);
            },
            property => {
                Assert.Equal("External", property.Name);
                Assert.Equal("Sheet1!A1", property.LinkedValue);
            });
        Assert.Equal("Body", Assert.Single(read.Document.Paragraphs).ToPlainText());
    }

    [Fact]
    public void Write_And_Read_Generator_Metadata() {
        RtfDocument document = RtfDocument.Create();
        document.Info.Generator = "Contoso RTF Writer";
        document.AddParagraph("Body");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = true });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"{\*\generator Contoso RTF Writer;}", rtf, StringComparison.Ordinal);
        Assert.Equal("Contoso RTF Writer", read.Document.Info.Generator);
        Assert.Equal("Body", Assert.Single(read.Document.Paragraphs).ToPlainText());
    }

    [Fact]
    public void Write_And_Read_Document_Variables() {
        RtfDocument document = RtfDocument.Create();
        document.AddDocumentVariable("Client", "Contoso");
        document.AddDocumentVariable("Region", "EMEA");
        document.AddParagraph("Body");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"{\*\docvar {Client}{Contoso}}{\*\docvar {Region}{EMEA}}", rtf, StringComparison.Ordinal);
        Assert.Collection(read.Document.DocumentVariables,
            variable => {
                Assert.Equal("Client", variable.Name);
                Assert.Equal("Contoso", variable.Value);
            },
            variable => {
                Assert.Equal("Region", variable.Name);
                Assert.Equal("EMEA", variable.Value);
            });
        Assert.Equal("Body", Assert.Single(read.Document.Paragraphs).ToPlainText());
    }
}
