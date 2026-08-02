using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class HtmlWordToHtml {
        [Fact]
        public void Test_WordToHtml_ComboBox_RoundTripsSelectedInternalValueWithDuplicateLabels() {
            using WordDocument document = WordDocument.Create();
            WordComboBox comboBox = document.AddParagraph().AddComboBox(new[] { "placeholder" });
            SdtContentComboBox properties = comboBox._sdtRun.SdtProperties!
                .GetFirstChild<SdtContentComboBox>()!;
            properties.RemoveAllChildren<ListItem>();
            properties.Append(
                new ListItem { Value = "id-a", DisplayText = "Same label" },
                new ListItem { Value = "id-b", DisplayText = "Same label" });
            properties.LastValue = "id-b";

            string html = document.ToHtml();

            Assert.Contains("value=\"Same label\" data-word-value=\"id-b\"", html,
                StringComparison.OrdinalIgnoreCase);
            using WordDocument roundTrip = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument();
            WordComboBox imported = Assert.Single(roundTrip.ComboBoxes);
            ListItem[] importedItems = imported._sdtRun.SdtProperties!
                .GetFirstChild<SdtContentComboBox>()!
                .Elements<ListItem>()
                .ToArray();
            Assert.Equal(new[] { "id-a", "id-b" }, importedItems.Select(item => item.Value?.Value).ToArray());
            Assert.All(importedItems, item => Assert.Equal("Same label", item.DisplayText?.Value));
            Assert.Equal("id-b", imported.SelectedValue);
        }
    }
}
