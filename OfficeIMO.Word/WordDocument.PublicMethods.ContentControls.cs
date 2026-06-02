using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Net.Http;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        /// <summary>
        /// Adds a new paragraph with a content control (structured document tag).
        /// </summary>
        /// <param name="text">Initial text of the control.</param>
        /// <param name="alias">Optional alias for the control.</param>
        /// <param name="tag">Optional tag for the control.</param>
        /// <returns>The created <see cref="WordStructuredDocumentTag"/>.</returns>
        public WordStructuredDocumentTag AddStructuredDocumentTag(string text, string? alias = null, string? tag = null) {
            return this.AddParagraph().AddStructuredDocumentTag(text, alias!, tag!);
        }

        /// <summary>
        /// Adds a new paragraph with a repeating section content control.
        /// </summary>
        /// <param name="sectionTitle">Optional title of the repeating section.</param>
        /// <param name="alias">Optional alias for the control.</param>
        /// <param name="tag">Optional tag for the control.</param>
        /// <returns>The created <see cref="WordRepeatingSection"/>.</returns>
        public WordRepeatingSection AddRepeatingSection(string? sectionTitle = null, string? alias = null, string? tag = null) {
            return this.AddParagraph().AddRepeatingSection(sectionTitle!, alias!, tag!);
        }

        /// <summary>
        /// Embeds another document as an alternative format part.
        /// </summary>
        /// <param name="fileName">Path to the document.</param>
        /// <param name="type">Optional format part type.</param>
        /// <returns>The created <see cref="WordEmbeddedDocument"/>.</returns>
        public WordEmbeddedDocument AddEmbeddedDocument(string fileName, WordAlternativeFormatImportPartType? type = null) {
            return new WordEmbeddedDocument(this, fileName, type, false);
        }

        /// <summary>
        /// Embeds HTML content as an alternative format part.
        /// </summary>
        /// <param name="htmlContent">HTML content to embed.</param>
        /// <param name="type">Format part type.</param>
        /// <returns>The created <see cref="WordEmbeddedDocument"/>.</returns>
        public WordEmbeddedDocument AddEmbeddedFragment(string htmlContent, WordAlternativeFormatImportPartType type) {
            return new WordEmbeddedDocument(this, htmlContent, type, true);
        }

        /// <summary>
        /// Retrieves a structured document tag by its tag value.
        /// </summary>
        /// <param name="tag">Tag value of the control.</param>
        /// <returns>The matching <see cref="WordStructuredDocumentTag"/> or <c>null</c>.</returns>
        public WordStructuredDocumentTag? GetStructuredDocumentTagByTag(string tag) {
            return this.StructuredDocumentTags.FirstOrDefault(sdt => sdt.Tag == tag);
        }

        /// <summary>
        /// Retrieves a structured document tag by its alias.
        /// </summary>
        /// <param name="alias">Alias of the control.</param>
        /// <returns>The matching <see cref="WordStructuredDocumentTag"/> or <c>null</c>.</returns>
        public WordStructuredDocumentTag? GetStructuredDocumentTagByAlias(string alias) {
            return this.StructuredDocumentTags.FirstOrDefault(sdt => sdt.Alias == alias);
        }

        /// <summary>
        /// Retrieves a checkbox control by its tag value.
        /// </summary>
        /// <param name="tag">Tag value of the checkbox.</param>
        /// <returns>The matching <see cref="WordCheckBox"/> or <c>null</c>.</returns>
        public WordCheckBox? GetCheckBoxByTag(string tag) {
            return this.CheckBoxes.FirstOrDefault(cb => cb.Tag == tag);
        }

        /// <summary>
        /// Retrieves a checkbox control by its alias.
        /// </summary>
        /// <param name="alias">Alias of the checkbox.</param>
        /// <returns>The matching <see cref="WordCheckBox"/> or <c>null</c>.</returns>
        public WordCheckBox? GetCheckBoxByAlias(string alias) {
            return this.CheckBoxes.FirstOrDefault(cb => cb.Alias == alias);
        }

        /// <summary>
        /// Retrieves a date picker control by its tag value.
        /// </summary>
        /// <param name="tag">Tag value of the date picker.</param>
        /// <returns>The matching <see cref="WordDatePicker"/> or <c>null</c>.</returns>
        public WordDatePicker? GetDatePickerByTag(string tag) {
            return this.DatePickers.FirstOrDefault(dp => dp.Tag == tag);
        }

        /// <summary>
        /// Retrieves a date picker control by its alias.
        /// </summary>
        /// <param name="alias">Alias of the date picker.</param>
        /// <returns>The matching <see cref="WordDatePicker"/> or <c>null</c>.</returns>
        public WordDatePicker? GetDatePickerByAlias(string alias) {
            return this.DatePickers.FirstOrDefault(dp => dp.Alias == alias);
        }

        /// <summary>
        /// Retrieves a dropdown list control by its tag value.
        /// </summary>
        /// <param name="tag">Tag value of the dropdown list.</param>
        /// <returns>The matching <see cref="WordDropDownList"/> or <c>null</c>.</returns>
        public WordDropDownList? GetDropDownListByTag(string tag) {
            return this.DropDownLists.FirstOrDefault(dl => dl.Tag == tag);
        }

        /// <summary>
        /// Retrieves a dropdown list control by its alias.
        /// </summary>
        /// <param name="alias">Alias of the dropdown list.</param>
        /// <returns>The matching <see cref="WordDropDownList"/> or <c>null</c>.</returns>
        public WordDropDownList? GetDropDownListByAlias(string alias) {
            return this.DropDownLists.FirstOrDefault(dl => dl.Alias == alias);
        }

        /// <summary>
        /// Retrieves a combo box control by its tag value.
        /// </summary>
        public WordComboBox? GetComboBoxByTag(string tag) {
            return this.ComboBoxes.FirstOrDefault(cb => cb.Tag == tag);
        }

        /// <summary>
        /// Retrieves a combo box control by its alias.
        /// </summary>
        public WordComboBox? GetComboBoxByAlias(string alias) {
            return this.ComboBoxes.FirstOrDefault(cb => cb.Alias == alias);
        }

        /// <summary>
        /// Retrieves a picture control by its tag value.
        /// </summary>
        public WordPictureControl? GetPictureControlByTag(string tag) {
            return this.PictureControls.FirstOrDefault(pc => pc.Tag == tag);
        }

        /// <summary>
        /// Retrieves a picture control by its alias.
        /// </summary>
        public WordPictureControl? GetPictureControlByAlias(string alias) {
            return this.PictureControls.FirstOrDefault(pc => pc.Alias == alias);
        }

        /// <summary>
        /// Retrieves a repeating section control by its tag value.
        /// </summary>
        public WordRepeatingSection? GetRepeatingSectionByTag(string tag) {
            return this.RepeatingSections.FirstOrDefault(rs => rs.Tag == tag);
        }

        /// <summary>
        /// Retrieves a repeating section control by its alias.
        /// </summary>
        public WordRepeatingSection? GetRepeatingSectionByAlias(string alias) {
            return this.RepeatingSections.FirstOrDefault(rs => rs.Alias == alias);
        }
        /// <summary>
        /// Removes an embedded document from the document.
        /// </summary>
        /// <param name="embeddedDocument">Embedded document to remove.</param>
        public void RemoveEmbeddedDocument(WordEmbeddedDocument embeddedDocument) {
            if (embeddedDocument == null) {
                throw new ArgumentNullException(nameof(embeddedDocument));
            }

            embeddedDocument.Remove();
        }
    }
}
