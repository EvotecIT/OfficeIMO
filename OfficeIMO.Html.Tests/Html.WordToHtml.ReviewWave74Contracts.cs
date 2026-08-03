using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class HtmlWordToHtml {
        [Fact]
        public void Test_WordToHtml_ComboBox_ExportsCanonicalSelectedInternalValue() {
            using WordDocument document = WordDocument.Create();
            WordComboBox comboBox = document.AddParagraph().AddComboBox(new[] { "placeholder" });
            SdtContentComboBox properties = comboBox._sdtRun.SdtProperties!
                .GetFirstChild<SdtContentComboBox>()!;
            ListItem item = Assert.Single(properties.Elements<ListItem>());
            item.Value = "ID-A";
            item.DisplayText = "Visible label";
            comboBox.SelectedValue = "id-a";

            string html = document.ToHtml();

            Assert.Contains("value=\"Visible label\" data-word-value=\"ID-A\"", html,
                StringComparison.OrdinalIgnoreCase);
            using WordDocument roundTrip = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument();
            WordComboBox imported = Assert.Single(roundTrip.ComboBoxes);
            Assert.Equal("ID-A", imported.SelectedValue);
            Assert.Equal(new[] { "Visible label" }, imported.Items.ToArray());
        }
    }

    public partial class Html {
        [Fact]
        public void HtmlToWord_ComboBox_MatchesInternalValueCaseInsensitively() {
            const string html = "<p><input type=\"text\" list=\"word-combo-1\" value=\"Visible label\" data-word-value=\"id-a\"><datalist id=\"word-combo-1\"><option value=\"Visible label\" data-word-value=\"ID-A\"></option></datalist></p>";

            using WordDocument document = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument();

            WordComboBox comboBox = Assert.Single(document.ComboBoxes);
            Assert.Equal("ID-A", comboBox.SelectedValue);
            ListItem item = Assert.Single(comboBox._sdtRun.SdtProperties!
                .GetFirstChild<SdtContentComboBox>()!
                .Elements<ListItem>());
            Assert.Equal("ID-A", item.Value?.Value);
            Assert.Equal("Visible label", item.DisplayText?.Value);
        }
    }
}
