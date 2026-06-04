using OfficeIMO.Word.Html;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_TextInput_BecomesStructuredDocumentTag() {
            const string html = "<p>Client <input type=\"text\" id=\"client-name\" name=\"client\" aria-label=\"Client name\" value=\"Contoso\"> approved</p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());

            var control = Assert.Single(doc.StructuredDocumentTags);
            Assert.Equal("Contoso", control.Text);
            Assert.Equal("Client name", control.Alias);
            Assert.Equal("client-name", control.Tag);
        }

        [Fact]
        public void HtmlToWord_Select_BecomesDropDownList() {
            const string html = "<p>Priority <select id=\"priority\" name=\"priority\" aria-label=\"Priority\"><option>Low</option><option selected>High</option></select> today</p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());

            var dropDown = Assert.Single(doc.DropDownLists);
            Assert.Equal(new[] { "Low", "High" }, dropDown.Items.ToArray());
            Assert.Equal("High", dropDown.SelectedValue);
            Assert.Equal("Priority", dropDown.Alias);
            Assert.Equal("priority", dropDown.Tag);
        }

        [Fact]
        public void HtmlToWord_Select_PreservesBlankOptions() {
            const string html = "<p>Status <select data-tag=\"status\"><option selected></option><option>Ready</option></select></p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());

            var dropDown = Assert.Single(doc.DropDownLists);
            Assert.Equal(new[] { string.Empty, "Ready" }, dropDown.Items.ToArray());
            Assert.Equal(string.Empty, dropDown.SelectedValue);
            Assert.Equal("status", dropDown.Tag);
        }

        [Fact]
        public void HtmlToWord_Select_ImportsOptGroupOptions() {
            const string html = "<p>Region <select data-tag=\"region\"><optgroup label=\"Europe\"><option>Poland</option><option selected>Germany</option></optgroup><option>Global</option></select></p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());

            var dropDown = Assert.Single(doc.DropDownLists);
            Assert.Equal(new[] { "Poland", "Germany", "Global" }, dropDown.Items.ToArray());
            Assert.Equal("Germany", dropDown.SelectedValue);
            Assert.Equal("region", dropDown.Tag);
        }

        [Fact]
        public void HtmlToWord_TextInputWithDatalist_BecomesComboBox() {
            const string html = "<p>Contact <input type=\"text\" list=\"word-combo-1\" data-tag=\"contact-method\" aria-label=\"Contact method\" value=\"Phone\"><datalist id=\"word-combo-1\"><option value=\"Email\"></option><option value=\"Phone\"></option></datalist></p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());

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

            var doc = html.LoadFromHtml(new HtmlToWordOptions());

            var comboBox = Assert.Single(doc.ComboBoxes);
            Assert.Equal(new[] { "Ready", string.Empty }, comboBox.Items.ToArray());
            Assert.Equal(string.Empty, comboBox.SelectedValue);
            Assert.Equal("status", comboBox.Tag);
        }

        [Fact]
        public void HtmlToWord_TextInputWithDatalist_DoesNotAddSpaceForSkippedMetadata() {
            const string html = "<p><input type=\"text\" list=\"word-combo-1\" value=\"Phone\"><datalist id=\"word-combo-1\"><option value=\"Email\"></option><option value=\"Phone\"></option></datalist></p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());

            var comboBox = Assert.Single(doc.ComboBoxes);
            Assert.Equal("Phone", comboBox.SelectedValue);
            var textRuns = doc._document.MainDocumentPart!.Document.Body!.Descendants<Text>().Select(text => text.Text).ToList();
            Assert.DoesNotContain(" ", textRuns);
        }

        [Fact]
        public void HtmlToWord_TextArea_BecomesStructuredDocumentTag() {
            const string html = "<p>Notes <textarea id=\"notes\" title=\"Review notes\">Line one\r\nLine two</textarea></p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());

            var control = Assert.Single(doc.StructuredDocumentTags);
            Assert.Equal("Line one\nLine two", control.Text);
            Assert.Equal("Review notes", control.Alias);
            Assert.Equal("notes", control.Tag);
        }

        [Fact]
        public void HtmlToWord_FormControls_ImportExportedDataTag() {
            const string html = "<p><input type=\"text\" data-tag=\"client-name\" aria-label=\"Client name\" value=\"Contoso\"><select data-tag=\"priority\" aria-label=\"Priority\"><option selected>High</option></select></p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());

            var control = Assert.Single(doc.StructuredDocumentTags, tag => tag.Text == "Contoso");
            Assert.Equal("client-name", control.Tag);
            var dropDown = Assert.Single(doc.DropDownLists);
            Assert.Equal("priority", dropDown.Tag);
        }

        [Fact]
        public void HtmlToWord_DateInput_BecomesDatePicker() {
            const string html = "<p>Due <input type=\"date\" data-tag=\"due-date\" aria-label=\"Due date\" value=\"2026-07-14\"></p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());

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
