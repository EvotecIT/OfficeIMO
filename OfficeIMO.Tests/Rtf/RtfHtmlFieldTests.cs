using OfficeIMO.Rtf;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlFieldTests {
    [Fact]
    public void Html_ToRtfDocument_Parses_Field_Metadata_In_Inline_Order() {
        const string html = "<p>Page <span data-officeimo-rtf-field=\"true\" data-officeimo-rtf-field-instruction=\"PAGE \\* MERGEFORMAT\"><strong>1</strong></span> done</p>";

        RtfDocument document = html.LoadRtfFromHtml();
        RtfParagraph paragraph = Assert.Single(document.Paragraphs);

        Assert.Collection(paragraph.Inlines,
            inline => Assert.Equal("Page ", Assert.IsType<RtfRun>(inline).Text),
            inline => {
                RtfField field = Assert.IsType<RtfField>(inline);
                Assert.Equal(@"PAGE \* MERGEFORMAT", field.Instruction);
                Assert.Equal("1", field.ToPlainText());
                Assert.Contains(field.Result.Runs, run => run.Text == "1" && run.Bold);
            },
            inline => Assert.Equal(" done", Assert.IsType<RtfRun>(inline).Text));

        string rtf = document.ToRtf();
        Assert.Contains(@"{\field{\*\fldinst PAGE \\* MERGEFORMAT}{\fldrslt \b 1\b0 }}", rtf, StringComparison.Ordinal);
    }

    [Fact]
    public void RtfDocument_ToHtml_Renders_Field_Metadata() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph("Page ");
        RtfField field = paragraph.AddField(@"PAGE \* MERGEFORMAT");
        field.AddText("1").SetBold();
        paragraph.AddText(" done");

        string html = document.ToHtml();

        Assert.Equal("<p>Page <span data-officeimo-rtf-field=\"true\" data-officeimo-rtf-field-instruction=\"PAGE \\* MERGEFORMAT\"><strong>1</strong></span> done</p>", html);

        RtfField roundTripField = Assert.IsType<RtfField>(Assert.Single(html.LoadRtfFromHtml().Paragraphs).Inlines[1]);
        Assert.Equal(@"PAGE \* MERGEFORMAT", roundTripField.Instruction);
        Assert.Equal("1", roundTripField.ToPlainText());
        Assert.Contains(roundTripField.Result.Runs, run => run.Text == "1" && run.Bold);
    }

    [Fact]
    public void RtfDocument_ToHtml_Escapes_Field_Instruction_Attribute() {
        RtfDocument document = RtfDocument.Create();
        RtfField field = document.AddParagraph().AddField("MERGEFIELD Patient<Name>");
        field.AddText("Ada");

        string html = document.ToHtml();

        Assert.Equal("<p><span data-officeimo-rtf-field=\"true\" data-officeimo-rtf-field-instruction=\"MERGEFIELD Patient&lt;Name&gt;\">Ada</span></p>", html);
        RtfField roundTripField = Assert.IsType<RtfField>(Assert.Single(html.LoadRtfFromHtml().Paragraphs).Inlines[0]);
        Assert.Equal("MERGEFIELD Patient<Name>", roundTripField.Instruction);
    }

    [Fact]
    public void Html_RoundTrips_Generated_Text_Metadata() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph("Page ");
        paragraph.AddPageNumber();
        paragraph.AddText(" Section ");
        paragraph.AddSectionNumber();
        paragraph.AddText(" Date ");
        paragraph.AddCurrentDate();
        paragraph.AddText(" Time ");
        paragraph.AddCurrentTime();

        string html = document.ToHtml();

        Assert.Equal("<p>Page <span data-officeimo-rtf-generated-text=\"page-number\"></span> Section <span data-officeimo-rtf-generated-text=\"section-number\"></span> Date <span data-officeimo-rtf-generated-text=\"current-date\"></span> Time <span data-officeimo-rtf-generated-text=\"current-time\"></span></p>", html);
        RtfParagraph roundTrip = Assert.Single(html.LoadRtfFromHtml().Paragraphs);
        Assert.Collection(roundTrip.Inlines,
            inline => Assert.Equal("Page ", Assert.IsType<RtfRun>(inline).Text),
            inline => Assert.Equal(RtfGeneratedTextKind.PageNumber, Assert.IsType<RtfGeneratedText>(inline).Kind),
            inline => Assert.Equal(" Section ", Assert.IsType<RtfRun>(inline).Text),
            inline => Assert.Equal(RtfGeneratedTextKind.SectionNumber, Assert.IsType<RtfGeneratedText>(inline).Kind),
            inline => Assert.Equal(" Date ", Assert.IsType<RtfRun>(inline).Text),
            inline => Assert.Equal(RtfGeneratedTextKind.CurrentDate, Assert.IsType<RtfGeneratedText>(inline).Kind),
            inline => Assert.Equal(" Time ", Assert.IsType<RtfRun>(inline).Text),
            inline => Assert.Equal(RtfGeneratedTextKind.CurrentTime, Assert.IsType<RtfGeneratedText>(inline).Kind));
    }

    [Fact]
    public void Html_ToRtfDocument_DoesNot_Preserve_Field_Span_As_Unknown_Text() {
        const string html = "<p><span data-officeimo-rtf-field=\"true\" data-officeimo-rtf-field-instruction=\"PAGE\">1</span></p>";

        RtfParagraph paragraph = Assert.Single(html.LoadRtfFromHtml(new RtfHtmlReadOptions { PreserveUnknownTagsAsText = true }).Paragraphs);

        RtfField field = Assert.IsType<RtfField>(Assert.Single(paragraph.Inlines));
        Assert.Equal("PAGE", field.Instruction);
        Assert.Equal("1", field.ToPlainText());
    }

    [Fact]
    public void Html_ToRtfDocument_Parses_Form_Field_Metadata() {
        const string html = "<p>Name: <span data-officeimo-rtf-field=\"true\" data-officeimo-rtf-field-instruction=\"FORMTEXT\" data-officeimo-rtf-form-field=\"true\" data-officeimo-rtf-form-controls=\"fftype=0;ffenabled=1;ffownhelp=1;ffownstat=1;ffprot=0;ffrecalc=1;ffmaxlen=50\" data-officeimo-rtf-form-name=\"Customer\" data-officeimo-rtf-form-default-text=\"Default\" data-officeimo-rtf-form-format=\"Uppercase\" data-officeimo-rtf-form-help-text=\"Help\" data-officeimo-rtf-form-status-text=\"Status\" data-officeimo-rtf-form-entry-macro=\"Enter\" data-officeimo-rtf-form-exit-macro=\"Exit\">Value</span></p>";

        RtfDocument document = html.LoadRtfFromHtml();
        RtfField field = Assert.IsType<RtfField>(Assert.Single(document.Paragraphs).Inlines[1]);

        Assert.Equal("FORMTEXT", field.Instruction);
        Assert.Equal("Value", field.ToPlainText());
        Assert.NotNull(field.FormFieldData);
        RtfFormFieldData data = field.FormFieldData!;
        Assert.Equal(RtfFormFieldKind.Text, data.Kind);
        Assert.True(data.Enabled);
        Assert.True(data.OwnHelp);
        Assert.True(data.OwnStatus);
        Assert.False(data.Protected);
        Assert.True(data.RecalculateOnExit);
        Assert.Equal(50, data.MaxLength);
        Assert.Equal("Customer", data.Name);
        Assert.Equal("Default", data.DefaultText);
        Assert.Equal("Uppercase", data.Format);
        Assert.Equal("Help", data.HelpText);
        Assert.Equal("Status", data.StatusText);
        Assert.Equal("Enter", data.EntryMacro);
        Assert.Equal("Exit", data.ExitMacro);

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        Assert.Contains(@"{\*\ffdata\fftype0\ffenabled1\ffownhelp1\ffownstat1\ffprot0\ffrecalc1\ffmaxlen50{\ffname Customer}{\ffdeftext Default}{\ffformat Uppercase}{\ffhelptext Help}{\ffstattext Status}{\ffentrymcr Enter}{\ffexitmcr Exit}}", rtf, StringComparison.Ordinal);
    }

    [Fact]
    public void RtfDocument_ToHtml_Renders_Form_Field_Metadata() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph("Choice: ");
        RtfField field = paragraph.AddField("FORMDROPDOWN");
        field.AddText("Second");
        field.SetFormFieldData(data => {
            data.Kind = RtfFormFieldKind.DropDown;
            data.Name = "Choice";
            data.Enabled = true;
            data.DefaultResult = 0;
            data.Result = 1;
            data.AddDropDownItem("First");
            data.AddDropDownItem("Second");
        });

        string html = document.ToHtml();

        Assert.Equal("<p>Choice: <span data-officeimo-rtf-field=\"true\" data-officeimo-rtf-field-instruction=\"FORMDROPDOWN\" data-officeimo-rtf-form-field=\"true\" data-officeimo-rtf-form-controls=\"fftype=2;ffenabled=1;ffdefres=0;ffres=1\" data-officeimo-rtf-form-name=\"Choice\" data-officeimo-rtf-form-dropdown-items=\"Rmlyc3Q=;U2Vjb25k\">Second</span></p>", html);

        RtfField roundTripField = Assert.IsType<RtfField>(Assert.Single(html.LoadRtfFromHtml().Paragraphs).Inlines[1]);
        Assert.NotNull(roundTripField.FormFieldData);
        RtfFormFieldData roundTripData = roundTripField.FormFieldData!;
        Assert.Equal(RtfFormFieldKind.DropDown, roundTripData.Kind);
        Assert.Equal("Choice", roundTripData.Name);
        Assert.True(roundTripData.Enabled);
        Assert.Equal(0, roundTripData.DefaultResult);
        Assert.Equal(1, roundTripData.Result);
        Assert.Equal(new[] { "First", "Second" }, roundTripData.DropDownItems);
        Assert.Equal("Second", roundTripField.ToPlainText());
    }

    [Fact]
    public void Html_ToRtfDocument_Ignores_Invalid_Form_Field_Control_Names() {
        const string html = "<p><span data-officeimo-rtf-field=\"true\" data-officeimo-rtf-field-instruction=\"FORMTEXT\" data-officeimo-rtf-form-controls=\"fftype=0;bad-name=1;ffmaxlen=20\">Value</span></p>";

        RtfField field = Assert.IsType<RtfField>(Assert.Single(html.LoadRtfFromHtml().Paragraphs).Inlines[0]);

        Assert.NotNull(field.FormFieldData);
        RtfFormFieldData data = field.FormFieldData!;
        Assert.Equal(RtfFormFieldKind.Text, data.Kind);
        Assert.Equal(20, data.MaxLength);
        Assert.DoesNotContain(data.Controls, control => control.Name == "bad-name");
        Assert.DoesNotContain("bad-name", field.Result.ToPlainText(), StringComparison.Ordinal);
    }
}
