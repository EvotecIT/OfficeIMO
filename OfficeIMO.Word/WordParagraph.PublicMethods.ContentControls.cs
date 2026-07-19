using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Reflection;
using System.Linq;
using System.Xml.Linq;
using MathParagraph = DocumentFormat.OpenXml.Math.Paragraph;
using OfficeMath = DocumentFormat.OpenXml.Math.OfficeMath;
using V = DocumentFormat.OpenXml.Vml;
using W14 = DocumentFormat.OpenXml.Office2010.Word;

namespace OfficeIMO.Word {
    /// <summary>
    /// Contains public methods for editing paragraphs.
    /// </summary>
    public partial class WordParagraph {
        /// <summary>
        /// Adds a simple content control (structured document tag) to the paragraph.
        /// </summary>
        /// <param name="text">Initial text of the control.</param>
        /// <param name="alias">Optional alias for the content control.</param>
        /// <param name="tag">Optional tag for the content control.</param>
        /// <returns>The created <see cref="WordStructuredDocumentTag"/> instance.</returns>
        public WordStructuredDocumentTag AddStructuredDocumentTag(string text = "", string? alias = null, string? tag = null) {
            var sdtRun = new SdtRun();

            var sdtProperties = new SdtProperties();
            if (!string.IsNullOrEmpty(alias)) {
                sdtProperties.Append(new SdtAlias() { Val = alias });
            }
            if (!string.IsNullOrEmpty(tag)) {
                sdtProperties.Append(new Tag() { Val = tag });
            }
            var sdtIdValue = _document.GenerateSdtId();
            sdtProperties.Append(new SdtId() { Val = new DocumentFormat.OpenXml.Int32Value(sdtIdValue) });

            var sdtContent = new SdtContentRun();
            var run = new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
            sdtContent.Append(run);

            sdtRun.Append(sdtProperties);
            sdtRun.Append(sdtContent);

            this._paragraph.Append(sdtRun);

            var paragraph = new WordParagraph(this._document, this._paragraph, sdtRun);
            return paragraph.StructuredDocumentTag!;
        }

        /// <summary>
        /// Adds a checkbox content control to the paragraph.
        /// </summary>
        /// <param name="isChecked">Initial checked state.</param>
        /// <param name="alias">Optional alias for the control.</param>
        /// <param name="tag">Optional tag for the control.</param>
        /// <returns>The created <see cref="WordCheckBox"/> instance.</returns>
        public WordCheckBox AddCheckBox(bool isChecked = false, string? alias = null, string? tag = null) {
            var sdtRun = new SdtRun();

            var props = new SdtProperties();
            if (!string.IsNullOrEmpty(alias)) {
                props.Append(new SdtAlias() { Val = alias });
            }
            if (!string.IsNullOrEmpty(tag)) {
                props.Append(new Tag() { Val = tag });
            }
            var sdtIdValue = _document.GenerateSdtId();
            props.Append(new SdtId() { Val = new DocumentFormat.OpenXml.Int32Value(sdtIdValue) });

            var checkBox = new W14.SdtContentCheckBox();
            checkBox.Append(new W14.Checked() { Val = isChecked ? W14.OnOffValues.One : W14.OnOffValues.Zero });
            checkBox.Append(new W14.CheckedState() {
                Font = WordCheckBox.SymbolFont,
                Val = WordCheckBox.CheckedStateValue
            });
            checkBox.Append(new W14.UncheckedState() {
                Font = WordCheckBox.SymbolFont,
                Val = WordCheckBox.UncheckedStateValue
            });
            props.Append(checkBox);

            var runProperties = new RunProperties();
            runProperties.Append(new RunFonts() {
                Ascii = WordCheckBox.SymbolFont,
                HighAnsi = WordCheckBox.SymbolFont,
                EastAsia = WordCheckBox.SymbolFont,
                ComplexScript = WordCheckBox.SymbolFont
            });

            var symbol = isChecked ? WordCheckBox.CheckedSymbol : WordCheckBox.UncheckedSymbol;
            var run = new Run(runProperties, new Text(symbol) {
                Space = SpaceProcessingModeValues.Preserve
            });

            var content = new SdtContentRun(run);

            sdtRun.Append(props);
            sdtRun.Append(content);

            this._paragraph.Append(sdtRun);

            var paragraph = new WordParagraph(this._document, this._paragraph, sdtRun);
            return paragraph.CheckBox!;
        }

        /// <summary>
        /// Removes the checkbox from the paragraph.
        /// </summary>
        public void RemoveCheckBox() {
            this.CheckBox?.Remove();
        }

        /// <summary>
        /// Sets the checked state of the paragraph's checkbox.
        /// </summary>
        /// <param name="value">New checked state.</param>
        public void SetCheckBoxValue(bool value) {
            if (this.CheckBox != null) {
                this.CheckBox.IsChecked = value;
            }
        }
        /// <summary>
        /// Adds a date picker content control to the paragraph.
        /// </summary>
        /// <param name="date">Initial date value.</param>
        /// <param name="alias">Optional alias for the control.</param>
        /// <param name="tag">Optional tag for the control.</param>
        /// <returns>The created <see cref="WordDatePicker"/> instance.</returns>
        public WordDatePicker AddDatePicker(System.DateTime? date = null, string? alias = null, string? tag = null) {
            var sdtRun = new SdtRun();

            var props = new SdtProperties();
            if (!string.IsNullOrEmpty(alias)) {
                props.Append(new SdtAlias() { Val = alias });
            }
            if (!string.IsNullOrEmpty(tag)) {
                props.Append(new Tag() { Val = tag });
            }
            var sdtIdValue = _document.GenerateSdtId();
            props.Append(new SdtId() { Val = new DocumentFormat.OpenXml.Int32Value(sdtIdValue) });

            var dateProp = new SdtContentDate();
            if (date.HasValue) {
                dateProp.FullDate = new DateTimeValue(date.Value);
            }
            props.Append(dateProp);

            var content = new SdtContentRun(new Run());

            sdtRun.Append(props);
            sdtRun.Append(content);

            this._paragraph.Append(sdtRun);

            return new WordDatePicker(this._document, this._paragraph, sdtRun);
        }

        /// <summary>
        /// Adds a dropdown list content control to the paragraph.
        /// </summary>
        /// <param name="items">Items to include in the list.</param>
        /// <param name="alias">Optional alias for the control.</param>
        /// <param name="tag">Optional tag for the control.</param>
        /// <returns>The created <see cref="WordDropDownList"/> instance.</returns>
        public WordDropDownList AddDropDownList(System.Collections.Generic.IEnumerable<string> items, string? alias = null, string? tag = null) {
            var sdtRun = new SdtRun();

            var props = new SdtProperties();
            if (!string.IsNullOrEmpty(alias)) {
                props.Append(new SdtAlias() { Val = alias });
            }
            if (!string.IsNullOrEmpty(tag)) {
                props.Append(new Tag() { Val = tag });
            }
            var sdtIdValue = _document.GenerateSdtId();
            props.Append(new SdtId() { Val = new DocumentFormat.OpenXml.Int32Value(sdtIdValue) });

            var ddl = new SdtContentDropDownList();
            if (items != null) {
                foreach (var item in items) {
                    ddl.Append(new ListItem() { DisplayText = item, Value = item });
                }
            }
            props.Append(ddl);

            var content = new SdtContentRun(new Run());

            sdtRun.Append(props);
            sdtRun.Append(content);

            this._paragraph.Append(sdtRun);

            return new WordDropDownList(this._document, this._paragraph, sdtRun);
        }

        /// <summary>
        /// Adds a combo box content control to the paragraph.
        /// </summary>
        /// <param name="items">Items to include in the combo box.</param>
        /// <param name="alias">Optional alias for the control.</param>
        /// <param name="tag">Optional tag for the control.</param>
        /// <param name="defaultValue">Optional default value to display; must match one of the provided items.</param>
        /// <returns>The created <see cref="WordComboBox"/> instance.</returns>
        public WordComboBox AddComboBox(System.Collections.Generic.IEnumerable<string> items, string? alias = null, string? tag = null, string? defaultValue = null) {
            var sdtRun = new SdtRun();

            var props = new SdtProperties();
            if (!string.IsNullOrEmpty(alias)) {
                props.Append(new SdtAlias() { Val = alias });
            }
            if (!string.IsNullOrEmpty(tag)) {
                props.Append(new Tag() { Val = tag });
            }
            var sdtIdValue = _document.GenerateSdtId();
            props.Append(new SdtId() { Val = new DocumentFormat.OpenXml.Int32Value(sdtIdValue) });

            var combo = new SdtContentComboBox();
            var itemList = items?.ToList() ?? new List<string>();
            if (defaultValue != null && itemList.All(item => item != defaultValue)) {
                throw new ArgumentException("The default combo box value must match one of the provided items.", nameof(defaultValue));
            }

            foreach (var item in itemList) {
                combo.Append(new ListItem() { DisplayText = item, Value = item });
            }
            props.Append(combo);

            string? selectedValue = defaultValue;
            if (string.IsNullOrEmpty(selectedValue) && itemList.Count > 0) {
                selectedValue = itemList[0];
            }

            if (!string.IsNullOrEmpty(selectedValue)) {
                combo.LastValue = selectedValue;
            }

            var run = new Run();
            if (!string.IsNullOrEmpty(selectedValue)) {
                run.Append(new Text(selectedValue!) { Space = SpaceProcessingModeValues.Preserve });
            }

            var content = new SdtContentRun(run);

            sdtRun.Append(props);
            sdtRun.Append(content);

            this._paragraph.Append(sdtRun);

            return new WordComboBox(this._document, this._paragraph, sdtRun);
        }

        /// <summary>
        /// Adds a picture content control containing an image to the paragraph.
        /// </summary>
        /// <param name="filePath">Image file path.</param>
        /// <param name="width">Optional width of the image.</param>
        /// <param name="height">Optional height of the image.</param>
        /// <param name="alias">Optional alias for the control.</param>
        /// <param name="tag">Optional tag for the control.</param>
        /// <returns>The created <see cref="WordPictureControl"/> instance.</returns>
        public WordPictureControl AddPictureControl(string filePath, double? width = null, double? height = null, string? alias = null, string? tag = null) {
            var sdtRun = new SdtRun();

            var props = new SdtProperties();
            if (!string.IsNullOrEmpty(alias)) {
                props.Append(new SdtAlias() { Val = alias });
            }
            if (!string.IsNullOrEmpty(tag)) {
                props.Append(new Tag() { Val = tag });
            }
            var sdtIdValue = _document.GenerateSdtId();
            props.Append(new SdtId() { Val = new DocumentFormat.OpenXml.Int32Value(sdtIdValue) });
            props.Append(new SdtContentPicture());

            var content = new SdtContentRun();
            var imageRun = new Run();
            var imageParagraph = new WordParagraph(this._document, this._paragraph, imageRun);
            imageParagraph.AddImage(filePath, width, height);
            content.Append(imageRun);

            sdtRun.Append(props);
            sdtRun.Append(content);

            this._paragraph.Append(sdtRun);

            return new WordPictureControl(this._document, this._paragraph, sdtRun);
        }

        /// <summary>
        /// Adds a repeating section content control to the paragraph.
        /// </summary>
        /// <param name="sectionTitle">Optional title of the repeating section.</param>
        /// <param name="alias">Optional alias for the control.</param>
        /// <param name="tag">Optional tag for the control.</param>
        /// <returns>The created <see cref="WordRepeatingSection"/> instance.</returns>
        public WordRepeatingSection AddRepeatingSection(string? sectionTitle = null, string? alias = null, string? tag = null) {
            var sdtIdValue = _document.GenerateSdtId();

            string xml = "<w:sdt xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' xmlns:w15='http://schemas.microsoft.com/office/word/2012/wordml'>";
            xml += "<w:sdtPr>";
            if (!string.IsNullOrEmpty(alias)) xml += $"<w:alias w:val='{alias}'/>";
            if (!string.IsNullOrEmpty(tag)) xml += $"<w:tag w:val='{tag}'/>";
            xml += "<w15:repeatingSection" + (string.IsNullOrEmpty(sectionTitle) ? string.Empty : $" w15:sectionTitle='{sectionTitle}'") + "/>";
            xml += "</w:sdtPr>";
            xml += "<w:sdtContent><w15:repeatingSectionItem><w:sdt><w:sdtContent><w:r/></w:sdtContent></w:sdt></w15:repeatingSectionItem></w:sdtContent>";
            xml += "</w:sdt>";

            var newSdt = new SdtRun(xml);

            // Repeating sections are still composed from raw XML because the Open XML SDK does not
            // expose strongly typed wrappers for the w15 repeating section vocabulary. Inject the
            // generated identifier through the object model to stay consistent with other helpers.
            var properties = newSdt.SdtProperties ?? new SdtProperties();
            properties.RemoveAllChildren<SdtId>();
            properties.Append(new SdtId { Val = new DocumentFormat.OpenXml.Int32Value(sdtIdValue) });
            newSdt.SdtProperties = properties;

            this._paragraph.Append(newSdt);

            return new WordRepeatingSection(this._document, this._paragraph, newSdt);
        }
    }
}
