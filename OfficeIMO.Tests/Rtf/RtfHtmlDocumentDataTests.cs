using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Html;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlDocumentDataTests {
    [Fact]
    public void RtfDocument_ToHtml_RoundTrips_User_Properties_And_Document_Variables() {
        RtfDocument document = RtfDocument.Create();
        document.AddUserProperty(RtfUserProperty.Text("Client", "Contoso"));
        document.AddUserProperty(RtfUserProperty.Boolean("Approved", true));
        document.AddUserProperty(RtfUserProperty.Number("Score", 98.5));
        RtfUserProperty linked = document.AddUserProperty("External");
        linked.LinkedValue = "Sheet1!A1";
        document.AddDocumentVariable("Mode", "Draft");
        document.AddDocumentVariable("Empty", string.Empty);
        document.AddParagraph("Body");

        string html = document.ToHtml(new RtfHtmlSaveOptions {
            FragmentOnly = false,
            NewLine = "\n"
        });

        Assert.Contains("<meta name=\"officeimo-rtf-user-properties\" content=\"", html, StringComparison.Ordinal);
        Assert.Contains("<meta name=\"officeimo-rtf-document-variables\" content=\"", html, StringComparison.Ordinal);

        RtfDocument roundTrip = html.LoadFromHtml();

        Assert.Collection(roundTrip.UserProperties,
            property => {
                Assert.Equal("Client", property.Name);
                Assert.Equal(RtfUserProperty.TextType, property.TypeCode);
                Assert.Equal("Contoso", property.StaticValue);
                Assert.Null(property.LinkedValue);
            },
            property => {
                Assert.Equal("Approved", property.Name);
                Assert.Equal(RtfUserProperty.BooleanType, property.TypeCode);
                Assert.Equal("1", property.StaticValue);
                Assert.Null(property.LinkedValue);
            },
            property => {
                Assert.Equal("Score", property.Name);
                Assert.Equal(RtfUserProperty.NumberType, property.TypeCode);
                Assert.Equal("98.5", property.StaticValue);
                Assert.Null(property.LinkedValue);
            },
            property => {
                Assert.Equal("External", property.Name);
                Assert.Null(property.TypeCode);
                Assert.Null(property.StaticValue);
                Assert.Equal("Sheet1!A1", property.LinkedValue);
            });

        Assert.Collection(roundTrip.DocumentVariables,
            variable => {
                Assert.Equal("Mode", variable.Name);
                Assert.Equal("Draft", variable.Value);
            },
            variable => {
                Assert.Equal("Empty", variable.Name);
                Assert.Equal(string.Empty, variable.Value);
            });

        string rtf = roundTrip.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        Assert.Contains(@"{\*\userprops{\propname Client}\proptype30{\staticval Contoso}{\propname Approved}\proptype11{\staticval 1}{\propname Score}\proptype5{\staticval 98.5}{\propname External}{\linkval Sheet1!A1}}", rtf, StringComparison.Ordinal);
        Assert.Contains(@"{\*\docvar {Mode}{Draft}}{\*\docvar {Empty}{}}", rtf, StringComparison.Ordinal);
    }
}
