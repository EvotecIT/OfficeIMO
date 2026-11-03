using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using System.Linq;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public partial class WordRtfConverterTests {
    [Fact]
    public void Word_Rtf_Bridge_Carries_Document_Metadata() {
        using WordDocument word = WordDocument.Create();
        var created = new DateTime(2026, 6, 15, 10, 20, 30);
        var modified = new DateTime(2026, 6, 16, 11, 21, 31);
        var printed = new DateTime(2026, 6, 17, 12, 22, 32);
        word.Properties.Title = "Bridge Title";
        word.Properties.Subject = "Bridge Subject";
        word.Properties.Creator = "Bridge Author";
        word.Properties.Category = "Bridge Category";
        word.Properties.Keywords = "rtf,word";
        word.Properties.Description = "Bridge Comments";
        word.Properties.LastModifiedBy = "Bridge Operator";
        word.Properties.Created = created;
        word.Properties.Modified = modified;
        word.Properties.LastPrinted = printed;
        word.ApplicationProperties.Company = "Bridge Company";
        word.ApplicationProperties.Manager = "Bridge Manager";
        word.ApplicationProperties.HyperlinkBase = "https://example.test/";
        word.ApplicationProperties.Pages = "9";
        word.ApplicationProperties.Characters = "123";
        word.ApplicationProperties.CharactersWithSpaces = "150";
        word.AddParagraph("Body");

        RtfDocument rtfDocument = word.ToRtfDocument();
        string rtf = word.ToRtf(new RtfSaveOptions { IncludeGenerator = false });
        using WordDocument roundTrip = RtfDocument.ParseResult(rtf).Document.ToWordDocument();

        Assert.Equal("Bridge Title", rtfDocument.Properties.Title);
        Assert.Equal("Bridge Subject", rtfDocument.Properties.Subject);
        Assert.Equal("Bridge Author", rtfDocument.Properties.Author);
        Assert.Equal("Bridge Company", rtfDocument.Properties.Company);
        Assert.Equal("Bridge Manager", rtfDocument.Properties.Manager);
        Assert.Equal("Bridge Operator", rtfDocument.Properties.Operator);
        Assert.Equal("Bridge Comments", rtfDocument.Properties.Comments);
        Assert.Equal("https://example.test/", rtfDocument.Properties.HyperlinkBase);
        Assert.Equal(created, rtfDocument.Properties.Created);
        Assert.Equal(modified, rtfDocument.Properties.Revised);
        Assert.Equal(printed, rtfDocument.Properties.Printed);
        Assert.Equal(9, rtfDocument.Properties.NumberOfPages);
        Assert.Equal(123, rtfDocument.Properties.NumberOfCharacters);
        Assert.Equal(150, rtfDocument.Properties.NumberOfCharactersWithSpaces);
        Assert.Contains(@"{\title Bridge Title}", rtf, StringComparison.Ordinal);
        Assert.Contains(@"{\creatim\yr2026\mo6\dy15\hr10\min20\sec30}", rtf, StringComparison.Ordinal);
        Assert.Equal("Bridge Title", roundTrip.Properties.Title);
        Assert.Equal("Bridge Author", roundTrip.Properties.Creator);
        Assert.Equal("Bridge Company", roundTrip.ApplicationProperties.Company);
        Assert.Equal("Bridge Manager", roundTrip.ApplicationProperties.Manager);
        Assert.Equal("Bridge Operator", roundTrip.Properties.LastModifiedBy);
        Assert.Equal(created, roundTrip.Properties.Created);
        Assert.Equal(modified, roundTrip.Properties.Modified);
        Assert.Equal(printed, roundTrip.Properties.LastPrinted);
        Assert.Equal("9", roundTrip.ApplicationProperties.Pages);
        Assert.Equal("123", roundTrip.ApplicationProperties.Characters);
        Assert.Equal("150", roundTrip.ApplicationProperties.CharactersWithSpaces);
    }

    [Fact]
    public void Word_Rtf_Bridge_Carries_Custom_Properties_And_Document_Variables() {
        using WordDocument word = WordDocument.Create();
        var due = new DateTime(2026, 6, 18, 13, 14, 15);
        word.CustomProperties["Client"] = new WordCustomProperty("Contoso");
        word.CustomProperties["Approved"] = new WordCustomProperty(true);
        word.CustomProperties["Ticket"] = new WordCustomProperty(42);
        word.CustomProperties["Score"] = new WordCustomProperty(98.5);
        word.CustomProperties["Due"] = new WordCustomProperty(due);
        word.SetDocumentVariable("Region", "EMEA");
        word.SetDocumentVariable("Mode", "Draft");
        word.AddParagraph("Body");

        RtfDocument rtfDocument = word.ToRtfDocument();
        string rtf = word.ToRtf(new RtfSaveOptions { IncludeGenerator = false });
        using WordDocument roundTrip = RtfDocument.ParseResult(rtf).Document.ToWordDocument();

        Assert.Collection(rtfDocument.UserProperties,
            property => {
                Assert.Equal("Approved", property.Name);
                Assert.Equal(RtfUserProperty.BooleanType, property.TypeCode);
                Assert.Equal("1", property.StaticValue);
            },
            property => {
                Assert.Equal("Client", property.Name);
                Assert.Equal(RtfUserProperty.TextType, property.TypeCode);
                Assert.Equal("Contoso", property.StaticValue);
            },
            property => {
                Assert.Equal("Due", property.Name);
                Assert.Equal(RtfUserProperty.DateTimeType, property.TypeCode);
                Assert.Equal("2026-06-18T13:14:15.0000000", property.StaticValue);
            },
            property => {
                Assert.Equal("Score", property.Name);
                Assert.Equal(RtfUserProperty.NumberType, property.TypeCode);
                Assert.Equal("98.5", property.StaticValue);
            },
            property => {
                Assert.Equal("Ticket", property.Name);
                Assert.Equal(RtfUserProperty.IntegerType, property.TypeCode);
                Assert.Equal("42", property.StaticValue);
            });
        Assert.Collection(rtfDocument.DocumentVariables,
            variable => {
                Assert.Equal("Mode", variable.Name);
                Assert.Equal("Draft", variable.Value);
            },
            variable => {
                Assert.Equal("Region", variable.Name);
                Assert.Equal("EMEA", variable.Value);
            });
        Assert.Contains(@"{\*\userprops{\propname Approved}\proptype11{\staticval 1}{\propname Client}\proptype30{\staticval Contoso}{\propname Due}\proptype64{\staticval 2026-06-18T13:14:15.0000000}{\propname Score}\proptype5{\staticval 98.5}{\propname Ticket}\proptype3{\staticval 42}}", rtf, StringComparison.Ordinal);
        Assert.Contains(@"{\*\docvar {Mode}{Draft}}{\*\docvar {Region}{EMEA}}", rtf, StringComparison.Ordinal);
        Assert.Equal("Contoso", roundTrip.CustomProperties["Client"].Text);
        Assert.Equal(true, roundTrip.CustomProperties["Approved"].Bool);
        Assert.Equal(42, roundTrip.CustomProperties["Ticket"].NumberInteger);
        Assert.Equal(98.5, roundTrip.CustomProperties["Score"].NumberDouble);
        Assert.Equal(due, roundTrip.CustomProperties["Due"].Date);
        Assert.Equal("Draft", roundTrip.GetDocumentVariable("Mode"));
        Assert.Equal("EMEA", roundTrip.GetDocumentVariable("Region"));
    }

    [Fact]
    public void Rtf_To_Word_Bridge_Carries_Linked_User_Property_Value_As_Text() {
        RtfDocument rtfDocument = RtfDocument.Create();
        RtfUserProperty property = rtfDocument.AddUserProperty("External");
        property.LinkedValue = "Sheet1!A1";
        rtfDocument.AddParagraph("Body");

        using WordDocument word = rtfDocument.ToWordDocument();

        WordCustomProperty customProperty = word.CustomProperties["External"];
        Assert.Equal(WordCustomPropertyType.Text, customProperty.PropertyType);
        Assert.Equal("Sheet1!A1", customProperty.Text);
    }
}
