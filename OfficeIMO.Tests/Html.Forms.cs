using OfficeIMO.Word.Html;
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
        }
    }
}
