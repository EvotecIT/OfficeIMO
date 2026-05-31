using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a repeating section content control within a paragraph.
    /// </summary>
    public class WordRepeatingSection : WordElement {
        private readonly WordDocument _document;
        private readonly Paragraph _paragraph;
        internal readonly SdtRun _sdtRun;

        internal WordRepeatingSection(WordDocument document, Paragraph paragraph, SdtRun sdtRun) {
            _document = document;
            _paragraph = paragraph;
            _sdtRun = sdtRun;
        }

        /// <summary>
        /// Gets the alias associated with this repeating section control.
        /// </summary>
        public string? Alias {
            get {
                var properties = _sdtRun.SdtProperties;
                var sdtAlias = properties?.OfType<SdtAlias>().FirstOrDefault();
                return sdtAlias?.Val;
            }
        }

        /// <summary>
        /// Gets or sets the tag value for this repeating section control.
        /// </summary>
        public string? Tag {
            get {
                var properties = _sdtRun.SdtProperties;
                var tag = properties?.OfType<Tag>().FirstOrDefault();
                return tag?.Val;
            }
            set {
                var properties = EnsureProperties();
                var tag = properties.OfType<Tag>().FirstOrDefault();
                if (value == null) {
                    tag?.Remove();
                    return;
                }
                if (tag == null) {
                    tag = new Tag();
                    properties.Append(tag);
                }
                tag.Val = value;
            }
        }

        /// <summary>
        /// Gets the text values in the repeating-section items.
        /// </summary>
        public IReadOnlyList<string> TextItems {
            get {
                return GetItems()
                    .Select(GetItemText)
                    .ToList();
            }
        }

        /// <summary>
        /// Replaces the repeating-section items with text values.
        /// </summary>
        /// <param name="values">Item text values to apply. An empty list leaves one blank item so the control stays editable in Word.</param>
        public void SetTextItems(IEnumerable<string?> values) {
            if (values == null) throw new ArgumentNullException(nameof(values));

            var textValues = values.Select(value => value ?? string.Empty).ToList();
            if (textValues.Count == 0) {
                textValues.Add(string.Empty);
            }

            SdtContentRun content = _sdtRun.SdtContentRun ??= new SdtContentRun();
            OpenXmlElement? template = GetItems().FirstOrDefault(item => item.Descendants<Text>().Any());

            foreach (OpenXmlElement existingItem in GetItems().ToList()) {
                existingItem.Remove();
            }

            foreach (string value in textValues) {
                OpenXmlElement item = template?.CloneNode(true) ?? CreateItem(string.Empty);
                SetItemText(item, value);
                content.Append(item);
            }
        }

        /// <summary>
        /// Extracts the repeating-section text values for content-control form maps.
        /// </summary>
        public IReadOnlyList<string> ExtractValue() {
            return TextItems;
        }

        /// <summary>
        /// Removes the repeating section from the paragraph.
        /// </summary>
        public void Remove() {
            _sdtRun.Remove();
        }

        private SdtProperties EnsureProperties() {
            if (_sdtRun.SdtProperties == null) {
                _sdtRun.SdtProperties = new SdtProperties();
            }
            return _sdtRun.SdtProperties;
        }

        private IEnumerable<OpenXmlElement> GetItems() {
            return _sdtRun.SdtContentRun?.ChildElements.Where(IsRepeatingSectionItem) ?? Enumerable.Empty<OpenXmlElement>();
        }

        private static OpenXmlElement CreateItem(string value) {
            var item = new OpenXmlUnknownElement("w15", "repeatingSectionItem", "http://schemas.microsoft.com/office/word/2012/wordml");
            item.Append(new SdtRun(
                new SdtContentRun(
                    new Run(
                        new Text(value) { Space = SpaceProcessingModeValues.Preserve }))));
            return item;
        }

        private static bool IsRepeatingSectionItem(OpenXmlElement element) {
            return element.LocalName == "repeatingSectionItem";
        }

        private static string GetItemText(OpenXmlElement item) {
            var typedText = item.Descendants<Text>().ToList();
            if (typedText.Count > 0) {
                return string.Concat(typedText.Select(text => text.Text));
            }

            if (!string.IsNullOrWhiteSpace(item.OuterXml)) {
                try {
                    var xml = System.Xml.Linq.XElement.Parse(item.OuterXml);
                    return string.Concat(xml.Descendants().Where(element => element.Name.LocalName == "t").Select(element => element.Value));
                } catch (System.Xml.XmlException) {
                    return string.Empty;
                }
            }

            return string.Empty;
        }

        private static void SetItemText(OpenXmlElement item, string value) {
            Text? text = item.Descendants<Text>().FirstOrDefault();
            if (text == null) {
                item.RemoveAllChildren();
                item.Append(new SdtRun(new SdtContentRun(new Run(new Text(value) { Space = SpaceProcessingModeValues.Preserve }))));
                return;
            }

            text.Space = SpaceProcessingModeValues.Preserve;
            text.Text = value;

            foreach (Text extraText in item.Descendants<Text>().Skip(1).ToList()) {
                extraText.Remove();
            }
        }
    }
}
