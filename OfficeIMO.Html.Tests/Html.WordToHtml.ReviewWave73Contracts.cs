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
            Assert.Equal("X", imported.SelectedValue);
        }

        [Fact]
        public void Test_WordToHtml_DropDownRoundTripsSelectedInternalValueWithDuplicateLabels() {
            const string html = "<select><option value='id-a'>Same label</option>"
                + "<option value='id-b' selected>Same label</option></select>";
            using WordDocument document =
                OfficeIMO.Html.HtmlConversionDocument.Parse(html)
                    .ToWordDocument();
            WordDropDownList dropDown = Assert.Single(document.DropDownLists);

            string exported = document.ToHtml();

            Assert.Equal("id-b", dropDown.SelectedValue);
            Assert.Contains("value=\"id-b\" selected", exported,
                StringComparison.OrdinalIgnoreCase);
            Assert.Equal(1, exported.Split(new[] { " selected" },
                StringSplitOptions.None).Length - 1);
            using WordDocument roundTrip =
                OfficeIMO.Html.HtmlConversionDocument.Parse(exported)
                    .ToWordDocument();
            Assert.Equal("id-b", Assert.Single(roundTrip.DropDownLists)
                .SelectedValue);
        }

        [Fact]
        public void Test_WordToHtml_DropDownPreservesSelectedEmptyInternalValueWithDuplicateLabels() {
            const string html = "<select><option value='' selected>Same label</option>"
                + "<option value='id-b'>Same label</option></select>";
            using WordDocument document =
                OfficeIMO.Html.HtmlConversionDocument.Parse(html)
                    .ToWordDocument();
            WordDropDownList dropDown = Assert.Single(document.DropDownLists);

            string exported = document.ToHtml();

            Assert.Equal(string.Empty, dropDown.SelectedValue);
            Assert.Contains("value=\"\" selected", exported,
                StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("value=\"id-b\" selected", exported,
                StringComparison.OrdinalIgnoreCase);
            Assert.Equal(1, exported.Split(new[] { " selected" },
                StringSplitOptions.None).Length - 1);
        }
    }
}
