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

        [Fact]
        public void Test_WordToHtml_DropDownSelectsOneDisplayMatchBeforeAnInternalValueMatch() {
            using WordDocument document = WordDocument.Create();
            WordDropDownList dropDown = document.AddParagraph().AddDropDownList(new[] { "placeholder" });
            SdtContentDropDownList properties = dropDown._sdtRun.SdtProperties!
                .GetFirstChild<SdtContentDropDownList>()!;
            properties.RemoveAllChildren<ListItem>();
            properties.Append(
                new ListItem { Value = "Red", DisplayText = "R" },
                new ListItem { Value = "X", DisplayText = "Red" });
            dropDown.SelectedValue = "Red";

            string html = document.ToHtml();

            Assert.DoesNotContain("value=\"Red\" selected", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("value=\"X\" selected", html, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(1, html.Split(new[] { " selected" }, StringSplitOptions.None).Length - 1);
            using WordDocument roundTrip = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument();
            WordDropDownList imported = Assert.Single(roundTrip.DropDownLists);
            ListItem[] importedItems = imported._sdtRun.SdtProperties!
                .GetFirstChild<SdtContentDropDownList>()!
                .Elements<ListItem>()
                .ToArray();
            Assert.Equal(new[] { "Red", "X" }, importedItems.Select(item => item.Value?.Value).ToArray());
            Assert.Equal("Red", imported.SelectedValue);
        }
    }
}
