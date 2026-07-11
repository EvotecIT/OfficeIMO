using OfficeIMO.Word.Html;
using OfficeIMO.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_TextInput_BecomesStructuredDocumentTag() {
            const string html = "<p>Client <input type=\"text\" id=\"client-name\" name=\"client\" aria-label=\"Client name\" value=\"Contoso\"> approved</p>";

            var doc = html.ToWordDocument(new HtmlToWordOptions());

            var control = Assert.Single(doc.StructuredDocumentTags);
            Assert.Equal("Contoso", control.Text);
            Assert.Equal("Client name", control.Alias);
            Assert.Equal("client-name", control.Tag);
        }

        [Theory]
        [InlineData("number", "42")]
        [InlineData("time", "14:30")]
        [InlineData("datetime-local", "2026-07-14T14:30")]
        [InlineData("month", "2026-07")]
        [InlineData("week", "2026-W29")]
        [InlineData("color", "#336699")]
        [InlineData("range", "75")]
        public void HtmlToWord_ValueInputTypes_BecomeStructuredDocumentTags(string type, string value) {
            string html = $"<p>Value <input type=\"{type}\" id=\"field\" name=\"field-name\" aria-label=\"Field\" value=\"{value}\"> saved</p>";

            var doc = html.ToWordDocument(new HtmlToWordOptions());

            var control = Assert.Single(doc.StructuredDocumentTags);
            Assert.Equal(value, control.Text);
            Assert.Equal("Field", control.Alias);
            Assert.Equal("field", control.Tag);
        }

        [Fact]
        public void HtmlToWord_NonDocumentInputControls_AreIgnored() {
            const string html = "<p>Start <input type=\"hidden\" value=\"secret\"><input type=\"file\" value=\"C:\\fakepath\\report.docx\"><input type=\"button\" value=\"Click\"><input type=\"submit\" value=\"Send\"> End</p>";

            var doc = html.ToWordDocument(new HtmlToWordOptions());

            Assert.Empty(doc.StructuredDocumentTags);
            Assert.Empty(doc.CheckBoxes);
            Assert.Empty(doc.DropDownLists);
            Assert.Empty(doc.ComboBoxes);
            var documentText = string.Concat(doc._document.MainDocumentPart!.Document.Body!.Descendants<Text>().Select(text => text.Text));
            Assert.DoesNotContain("secret", documentText, StringComparison.Ordinal);
            Assert.DoesNotContain("fakepath", documentText, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Click", documentText, StringComparison.Ordinal);
            Assert.DoesNotContain("Send", documentText, StringComparison.Ordinal);
        }

        [Fact]
        public void HtmlToWord_ProgressAndMeter_BecomeStructuredDocumentTags() {
            const string html = "<p>Build <progress id=\"build-progress\" aria-label=\"Build progress\" value=\"40\" max=\"100\"></progress> Quality <meter id=\"quality\" title=\"Quality score\" value=\"0.82\" max=\"1\"></meter></p>";

            var doc = html.ToWordDocument(new HtmlToWordOptions());

            Assert.Equal(2, doc.StructuredDocumentTags.Count);
            Assert.Contains(doc.StructuredDocumentTags, control =>
                control.Text == "40 / 100" &&
                control.Alias == "Build progress" &&
                control.Tag == "build-progress");
            Assert.Contains(doc.StructuredDocumentTags, control =>
                control.Text == "0.82 / 1" &&
                control.Alias == "Quality score" &&
                control.Tag == "quality");
        }

        [Fact]
        public void HtmlToWord_ProgressWithoutValue_UsesFallbackText() {
            const string html = "<p>Status <progress id=\"download\">Pending</progress></p>";

            var doc = html.ToWordDocument(new HtmlToWordOptions());

            var control = Assert.Single(doc.StructuredDocumentTags);
            Assert.Equal("Pending", control.Text);
            Assert.Equal("download", control.Tag);
        }

        [Fact]
        public void HtmlToWord_Select_BecomesDropDownList() {
            const string html = "<p>Priority <select id=\"priority\" name=\"priority\" aria-label=\"Priority\"><option>Low</option><option selected>High</option></select> today</p>";

            var doc = html.ToWordDocument(new HtmlToWordOptions());

            var dropDown = Assert.Single(doc.DropDownLists);
            Assert.Equal(new[] { "Low", "High" }, dropDown.Items.ToArray());
            Assert.Equal("High", dropDown.SelectedValue);
            Assert.Equal("Priority", dropDown.Alias);
            Assert.Equal("priority", dropDown.Tag);
        }

        [Fact]
        public void HtmlToWord_Select_PreservesBlankOptions() {
            const string html = "<p>Status <select data-tag=\"status\"><option selected></option><option>Ready</option></select></p>";

            var doc = html.ToWordDocument(new HtmlToWordOptions());

            var dropDown = Assert.Single(doc.DropDownLists);
            Assert.Equal(new[] { string.Empty, "Ready" }, dropDown.Items.ToArray());
            Assert.Equal(string.Empty, dropDown.SelectedValue);
            Assert.Equal("status", dropDown.Tag);
        }

        [Fact]
        public void HtmlToWord_Select_ImportsOptGroupOptions() {
            const string html = "<p>Region <select data-tag=\"region\"><optgroup label=\"Europe\"><option>Poland</option><option selected>Germany</option></optgroup><option>Global</option></select></p>";

            var doc = html.ToWordDocument(new HtmlToWordOptions());

            var dropDown = Assert.Single(doc.DropDownLists);
            Assert.Equal(new[] { "Poland", "Germany", "Global" }, dropDown.Items.ToArray());
            Assert.Equal("Germany", dropDown.SelectedValue);
            Assert.Equal("region", dropDown.Tag);
        }

        [Fact]
        public void HtmlToWord_MultiSelect_BecomesStructuredDocumentTagWithSelectedValues() {
            const string html = "<p>Regions <select multiple data-tag=\"regions\" aria-label=\"Regions\"><option selected>Poland</option><option>Germany</option><option value=\"global\" selected>Global</option></select></p>";

            var doc = html.ToWordDocument(new HtmlToWordOptions());

            Assert.Empty(doc.DropDownLists);
            var control = Assert.Single(doc.StructuredDocumentTags);
            Assert.Equal("Poland\nglobal", control.Text);
            Assert.Equal("Regions", control.Alias);
            Assert.Equal("regions", control.Tag);
        }

        [Fact]
        public void HtmlToWord_MultiSelectWithoutSelection_DoesNotDefaultToFirstOption() {
            const string html = "<p>Regions <select multiple data-tag=\"regions\"><option>Poland</option><option>Germany</option></select></p>";

            var doc = html.ToWordDocument(new HtmlToWordOptions());

            Assert.Empty(doc.DropDownLists);
            var control = Assert.Single(doc.StructuredDocumentTags);
            Assert.Equal(string.Empty, control.Text);
            Assert.Equal("regions", control.Tag);
        }

        [Fact]
        public void HtmlToWord_MultiSelect_SavesAsValidOpenXmlDocument() {
            const string html = "<p>Regions <select multiple data-tag=\"regions\" aria-label=\"Regions\"><option selected>Poland</option><option value=\"global\" selected>Global</option></select></p>";

            using var doc = html.ToWordDocument(new HtmlToWordOptions());

            using MemoryStream stream = doc.SaveAsMemoryStream();
            using WordprocessingDocument package = WordprocessingDocument.Open(stream, false);
            var errors = new OpenXmlValidator().Validate(package).ToList();
            Assert.True(errors.Count == 0, OpenXmlValidationFormatting.FormatValidationErrors(errors));
        }

        [Fact]
        public void HtmlToWord_RadioGroup_BecomesSingleDropDownList() {
            const string html = "<p>Priority <label><input type=\"radio\" name=\"priority\" value=\"low\"> Low</label><label><input type=\"radio\" name=\"priority\" value=\"high\" checked> High</label> today</p>";

            var doc = html.ToWordDocument(new HtmlToWordOptions());

            var dropDown = Assert.Single(doc.DropDownLists);
            Assert.Equal(new[] { "low", "high" }, dropDown.Items.ToArray());
            Assert.Equal("high", dropDown.SelectedValue);
            Assert.Equal("priority", dropDown.Alias);
            Assert.Equal("priority", dropDown.Tag);
            var visibleText = string.Concat(doc._document.MainDocumentPart!.Document.Body!.Descendants<Text>().Select(text => text.Text));
            Assert.DoesNotContain("Low", visibleText, StringComparison.Ordinal);
        }

        [Fact]
        public void HtmlToWord_RadioGroupWithoutValue_UsesLabelText() {
            const string html = "<p><label><input type=\"radio\" name=\"priority\"> Low</label><label><input type=\"radio\" name=\"priority\" checked> High</label></p>";

            var doc = html.ToWordDocument(new HtmlToWordOptions());

            var dropDown = Assert.Single(doc.DropDownLists);
            Assert.Equal(new[] { "Low", "High" }, dropDown.Items.ToArray());
            Assert.Equal("High", dropDown.SelectedValue);
        }

        [Fact]
        public void HtmlToWord_RadioGroupWithoutSelection_DoesNotDefaultToFirstOption() {
            const string html = "<p>Contact <input type=\"radio\" name=\"contact\" value=\"email\"><input type=\"radio\" name=\"contact\" value=\"phone\"></p>";

            var doc = html.ToWordDocument(new HtmlToWordOptions());

            var dropDown = Assert.Single(doc.DropDownLists);
            Assert.Equal(new[] { string.Empty, "email", "phone" }, dropDown.Items.ToArray());
            Assert.Equal(string.Empty, dropDown.SelectedValue);
        }

        [Fact]
        public void HtmlToWord_RadioGroupsWithSameNameInDifferentForms_StaySeparate() {
            const string html = "<form><input type=\"radio\" name=\"status\" value=\"internal\" checked></form><form><input type=\"radio\" name=\"status\" value=\"external\" checked></form>";

            var doc = html.ToWordDocument(new HtmlToWordOptions());

            Assert.Equal(2, doc.DropDownLists.Count);
            Assert.Contains(doc.DropDownLists, dropDown => dropDown.Items.SequenceEqual(new[] { "internal" }) && dropDown.SelectedValue == "internal");
            Assert.Contains(doc.DropDownLists, dropDown => dropDown.Items.SequenceEqual(new[] { "external" }) && dropDown.SelectedValue == "external");
        }

        [Fact]
        public void HtmlToWord_RadioGroupWithExplicitAndAncestorFormOwners_StaysTogether() {
            const string html = "<form id=\"f\"><input type=\"radio\" name=\"status\" value=\"internal\" checked></form><input type=\"radio\" name=\"status\" form=\"f\" value=\"external\">";

            var doc = html.ToWordDocument(new HtmlToWordOptions());

            var dropDown = Assert.Single(doc.DropDownLists);
            Assert.Equal(new[] { "internal", "external" }, dropDown.Items.ToArray());
            Assert.Equal("internal", dropDown.SelectedValue);
        }

        [Fact]
        public void HtmlToWord_RadioGroup_SavesAsValidOpenXmlDocument() {
            const string html = "<p>Priority <label><input type=\"radio\" name=\"priority\" value=\"low\"> Low</label><label><input type=\"radio\" name=\"priority\" value=\"high\" checked> High</label></p>";

            using var doc = html.ToWordDocument(new HtmlToWordOptions());

            using MemoryStream stream = doc.SaveAsMemoryStream();
            using WordprocessingDocument package = WordprocessingDocument.Open(stream, false);
            var errors = new OpenXmlValidator().Validate(package).ToList();
            Assert.True(errors.Count == 0, OpenXmlValidationFormatting.FormatValidationErrors(errors));
        }

        [Fact]
        public void HtmlToWord_TextInputWithDatalist_BecomesComboBox() {
            const string html = "<p>Contact <input type=\"text\" list=\"word-combo-1\" data-tag=\"contact-method\" aria-label=\"Contact method\" value=\"Phone\"><datalist id=\"word-combo-1\"><option value=\"Email\"></option><option value=\"Phone\"></option></datalist></p>";

            var doc = html.ToWordDocument(new HtmlToWordOptions());

            var comboBox = Assert.Single(doc.ComboBoxes);
            Assert.Equal(new[] { "Email", "Phone" }, comboBox.Items.ToArray());
            Assert.Equal("Phone", comboBox.SelectedValue);
            Assert.Equal("Contact method", comboBox.Alias);
            Assert.Equal("contact-method", comboBox.Tag);
            Assert.DoesNotContain(doc.Paragraphs, paragraph => paragraph.Text.Contains("Email", StringComparison.Ordinal));
        }

        [Fact]
        public void HtmlToWord_TextInputWithDatalist_PreservesBlankSelectedValue() {
            const string html = "<p>Status <input type=\"text\" list=\"word-combo-1\" data-tag=\"status\"><datalist id=\"word-combo-1\"><option value=\"Ready\"></option><option value=\"\"></option></datalist></p>";

            var doc = html.ToWordDocument(new HtmlToWordOptions());

            var comboBox = Assert.Single(doc.ComboBoxes);
            Assert.Equal(new[] { "Ready", string.Empty }, comboBox.Items.ToArray());
            Assert.Equal(string.Empty, comboBox.SelectedValue);
            Assert.Equal("status", comboBox.Tag);
        }

        [Fact]
        public void HtmlToWord_TextInputWithDatalist_DoesNotAddSpaceForSkippedMetadata() {
            const string html = "<p><input type=\"text\" list=\"word-combo-1\" value=\"Phone\"><datalist id=\"word-combo-1\"><option value=\"Email\"></option><option value=\"Phone\"></option></datalist></p>";

            var doc = html.ToWordDocument(new HtmlToWordOptions());

            var comboBox = Assert.Single(doc.ComboBoxes);
            Assert.Equal("Phone", comboBox.SelectedValue);
            var textRuns = doc._document.MainDocumentPart!.Document.Body!.Descendants<Text>().Select(text => text.Text).ToList();
            Assert.DoesNotContain(" ", textRuns);
        }

        [Fact]
        public void HtmlToWord_TextArea_BecomesStructuredDocumentTag() {
            const string html = "<p>Notes <textarea id=\"notes\" title=\"Review notes\">Line one\r\nLine two</textarea></p>";

            var doc = html.ToWordDocument(new HtmlToWordOptions());

            var control = Assert.Single(doc.StructuredDocumentTags);
            Assert.Equal("Line one\nLine two", control.Text);
            Assert.Equal("Review notes", control.Alias);
            Assert.Equal("notes", control.Tag);
        }

        [Fact]
        public void HtmlToWord_FormControls_ImportExportedDataTag() {
            const string html = "<p><input type=\"text\" data-tag=\"client-name\" aria-label=\"Client name\" value=\"Contoso\"><select data-tag=\"priority\" aria-label=\"Priority\"><option selected>High</option></select></p>";

            var doc = html.ToWordDocument(new HtmlToWordOptions());

            var control = Assert.Single(doc.StructuredDocumentTags, tag => tag.Text == "Contoso");
            Assert.Equal("client-name", control.Tag);
            var dropDown = Assert.Single(doc.DropDownLists);
            Assert.Equal("priority", dropDown.Tag);
        }

        [Fact]
        public void HtmlToWord_DateInput_BecomesDatePicker() {
            const string html = "<p>Due <input type=\"date\" data-tag=\"due-date\" aria-label=\"Due date\" value=\"2026-07-14\"></p>";

            var doc = html.ToWordDocument(new HtmlToWordOptions());

            var datePicker = Assert.Single(doc.DatePickers);
            Assert.Equal(new DateTime(2026, 7, 14), datePicker.Date);
            Assert.Equal("Due date", datePicker.Alias);
            Assert.Equal("due-date", datePicker.Tag);
            var displayedText = doc._document.MainDocumentPart!.Document.Body!.Descendants<SdtRun>()
                .Single(sdt => sdt.SdtProperties?.Elements<SdtContentDate>().Any() == true)
                .SdtContentRun!.Descendants<Text>().Single().Text;
            Assert.Equal("2026-07-14", displayedText);
        }
    }
}
