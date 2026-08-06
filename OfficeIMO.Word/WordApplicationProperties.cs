using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;

using System.Globalization;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides strongly typed access to the extended application properties
    /// stored in the underlying <see cref="WordprocessingDocument"/>.
    /// </summary>
    public class WordApplicationProperties {
        private readonly WordprocessingDocument _wordprocessingDocument;
        private readonly WordDocument _document;

        /// <summary>
        /// Gets or sets the application name that created the document.
        /// </summary>
        public string Application {
            get {
                var properties = _wordprocessingDocument.ExtendedFilePropertiesPart?.Properties;
                return properties?.Application?.Text ?? string.Empty;
            }
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }

                properties.Application ??= new Application();
                properties.Application.Text = value;
            }
        }
        /// <summary>
        /// Gets or sets the version of the application that created the document.
        /// </summary>
        public string ApplicationVersion {
            get {
                var properties = _wordprocessingDocument.ExtendedFilePropertiesPart?.Properties;
                return properties?.ApplicationVersion?.Text ?? string.Empty;
            }
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }

                properties.ApplicationVersion ??= new ApplicationVersion();
                properties.ApplicationVersion.Text = value;
            }
        }
        /// <summary>
        /// Gets or sets the total number of paragraphs in the document.
        /// </summary>
        public string Paragraphs {
            get {
                var properties = _wordprocessingDocument.ExtendedFilePropertiesPart?.Properties;
                return properties?.Paragraphs?.Text ?? string.Empty;
            }
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }

                properties.Paragraphs ??= new Paragraphs();
                properties.Paragraphs.Text = value;
            }
        }
        /// <summary>
        /// Gets or sets the total number of pages in the document.
        /// </summary>
        public string Pages {
            get {
                var properties = _wordprocessingDocument.ExtendedFilePropertiesPart?.Properties;
                return properties?.Pages?.Text ?? string.Empty;
            }
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }

                properties.Pages ??= new Pages();
                properties.Pages.Text = value;
            }
        }
        /// <summary>
        /// Gets or sets the character count of the document.
        /// </summary>
        public string Characters {
            get {
                var properties = _wordprocessingDocument.ExtendedFilePropertiesPart?.Properties;
                return properties?.Characters?.Text ?? string.Empty;
            }
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }

                properties.Characters ??= new Characters();
                properties.Characters.Text = value;
            }
        }
        /// <summary>
        /// Gets or sets the character count including spaces.
        /// </summary>
        public string CharactersWithSpaces {
            get {
                var properties = _wordprocessingDocument.ExtendedFilePropertiesPart?.Properties;
                return properties?.CharactersWithSpaces?.Text ?? string.Empty;
            }
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }

                properties.CharactersWithSpaces ??= new CharactersWithSpaces();
                properties.CharactersWithSpaces.Text = value;
            }
        }
        /// <summary>
        /// Gets or sets the company associated with the document.
        /// </summary>
        public string Company {
            get {
                var properties = _wordprocessingDocument.ExtendedFilePropertiesPart?.Properties;
                return properties?.Company?.Text ?? string.Empty;
            }
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }

                properties.Company ??= new Company();
                properties.Company.Text = value;
            }
        }
        /// <summary>
        /// Gets or sets whether the extended-properties part contains legacy digital-signature metadata.
        /// This flag does not indicate that the package has a valid cryptographic signature.
        /// </summary>
        public bool HasDigitalSignatureMetadata {
            get => _wordprocessingDocument.ExtendedFilePropertiesPart?.Properties?.DigitalSignature != null;
            set {
                var properties = GetProperties();
                if (properties != null) {
                    properties.DigitalSignature = value ? new DigitalSignature() : null;
                }
            }
        }
        /// <summary>
        /// Gets or sets the document security information.
        /// </summary>
        public int? DocumentSecurity {
            get => ReadInt32(_wordprocessingDocument.ExtendedFilePropertiesPart?.Properties?.DocumentSecurity);
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }
                properties.DocumentSecurity = value.HasValue
                    ? new DocumentSecurity { Text = value.Value.ToString(CultureInfo.InvariantCulture) }
                    : null;
            }
        }
        /// <summary>
        /// Gets or sets the heading pairs associated with the document.
        /// </summary>
        internal HeadingPairs? HeadingPairs {
            get => _wordprocessingDocument.ExtendedFilePropertiesPart?.Properties?.HeadingPairs;
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }
                properties.HeadingPairs = value;
            }
        }
        /// <summary>
        /// Gets or sets the hidden slides information for the document.
        /// </summary>
        public int? HiddenSlides {
            get => ReadInt32(_wordprocessingDocument.ExtendedFilePropertiesPart?.Properties?.HiddenSlides);
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }
                properties.HiddenSlides = value.HasValue
                    ? new HiddenSlides { Text = value.Value.ToString(CultureInfo.InvariantCulture) }
                    : null;
            }
        }
        /// <summary>
        /// Gets or sets the base address used for resolving hyperlinks.
        /// </summary>
        public string? HyperlinkBase {
            get => _wordprocessingDocument.ExtendedFilePropertiesPart?.Properties?.HyperlinkBase?.Text;
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }
                properties.HyperlinkBase = value == null ? null : new HyperlinkBase { Text = value };
            }
        }
        /// <summary>
        /// Gets or sets the list of hyperlinks in the document.
        /// </summary>
        internal HyperlinkList? HyperlinkList {
            get => _wordprocessingDocument.ExtendedFilePropertiesPart?.Properties?.HyperlinkList;
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }
                properties.HyperlinkList = value;
            }
        }
        /// <summary>
        /// Gets or sets the total number of lines in the document.
        /// </summary>
        public int? Lines {
            get => ReadInt32(_wordprocessingDocument.ExtendedFilePropertiesPart?.Properties?.Lines);
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }
                properties.Lines = value.HasValue
                    ? new Lines { Text = value.Value.ToString(CultureInfo.InvariantCulture) }
                    : null;
            }
        }
        /// <summary>
        /// Gets or sets the manager associated with the document.
        /// </summary>
        public string? Manager {
            get => _wordprocessingDocument.ExtendedFilePropertiesPart?.Properties?.Manager?.Text;
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }
                properties.Manager = value == null ? null : new Manager { Text = value };
            }
        }
        /// <summary>
        /// Gets or sets a value indicating whether hyperlinks have changed.
        /// </summary>
        public bool? HyperlinksChanged {
            get => ReadBoolean(_wordprocessingDocument.ExtendedFilePropertiesPart?.Properties?.HyperlinksChanged);
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }
                properties.HyperlinksChanged = value.HasValue
                    ? new HyperlinksChanged { Text = value.Value ? "true" : "false" }
                    : null;
            }
        }
        /// <summary>
        /// Gets or sets the notes information for the document.
        /// </summary>
        public int? Notes {
            get => ReadInt32(_wordprocessingDocument.ExtendedFilePropertiesPart?.Properties?.Notes);
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }
                properties.Notes = value.HasValue
                    ? new Notes { Text = value.Value.ToString(CultureInfo.InvariantCulture) }
                    : null;
            }
        }
        /// <summary>
        /// Gets or sets the multimedia clips associated with the document.
        /// </summary>
        public int? MultimediaClips {
            get => ReadInt32(_wordprocessingDocument.ExtendedFilePropertiesPart?.Properties?.MultimediaClips);
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }
                properties.MultimediaClips = value.HasValue
                    ? new MultimediaClips { Text = value.Value.ToString(CultureInfo.InvariantCulture) }
                    : null;
            }
        }
        /// <summary>
        /// Gets or sets the total editing time for the document.
        /// </summary>
        public int? TotalTime {
            get => ReadInt32(_wordprocessingDocument.ExtendedFilePropertiesPart?.Properties?.TotalTime);
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }
                properties.TotalTime = value.HasValue
                    ? new TotalTime { Text = value.Value.ToString(CultureInfo.InvariantCulture) }
                    : null;
            }
        }
        /// <summary>
        /// Gets or sets the scale crop information for the document.
        /// </summary>
        public bool? ScaleCrop {
            get => ReadBoolean(_wordprocessingDocument.ExtendedFilePropertiesPart?.Properties?.ScaleCrop);
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }
                properties.ScaleCrop = value.HasValue
                    ? new ScaleCrop { Text = value.Value ? "true" : "false" }
                    : null;
            }
        }
        /// <summary>
        /// Gets or sets the presentation format used by the document.
        /// </summary>
        public string? PresentationFormat {
            get => _wordprocessingDocument.ExtendedFilePropertiesPart?.Properties?.PresentationFormat?.Text;
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }
                properties.PresentationFormat = value == null ? null : new PresentationFormat { Text = value };
            }
        }
        /// <summary>
        /// Gets or sets the template from which the document was created.
        /// </summary>
        public string? Template {
            get => _wordprocessingDocument.ExtendedFilePropertiesPart?.Properties?.Template?.Text;
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }
                properties.Template = value == null ? null : new Template { Text = value };
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the document is shared.
        /// </summary>
        public bool? SharedDocument {
            get => ReadBoolean(_wordprocessingDocument.ExtendedFilePropertiesPart?.Properties?.SharedDocument);
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }
                properties.SharedDocument = value.HasValue
                    ? new SharedDocument { Text = value.Value ? "true" : "false" }
                    : null;
            }
        }

        /// <summary>
        /// Gets or sets the total number of words in the document.
        /// </summary>
        public int? Words {
            get => ReadInt32(_wordprocessingDocument.ExtendedFilePropertiesPart?.Properties?.Words);
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }
                properties.Words = value.HasValue
                    ? new Words { Text = value.Value.ToString(CultureInfo.InvariantCulture) }
                    : null;
            }
        }

        /// <summary>
        /// Initializes a new instance bound to the specified document.
        /// </summary>
        /// <param name="document">Parent document.</param>
        public WordApplicationProperties(WordDocument document) {
            _document = document;
            _wordprocessingDocument = document._wordprocessingDocument;
            _document.ApplicationProperties = this;
        }

        private Properties? GetProperties() {
            CreateExtendedFileProperties();
            return _wordprocessingDocument.ExtendedFilePropertiesPart?.Properties;
        }

        private static int? ReadInt32(DocumentFormat.OpenXml.OpenXmlLeafTextElement? value) {
            return int.TryParse(value?.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed)
                ? parsed
                : null;
        }

        private static bool? ReadBoolean(DocumentFormat.OpenXml.OpenXmlLeafTextElement? value) {
            string? text = value?.Text;
            if (string.Equals(text, "1", StringComparison.Ordinal)) return true;
            if (string.Equals(text, "0", StringComparison.Ordinal)) return false;
            return bool.TryParse(text, out bool parsed) ? parsed : null;
        }

        private void CreateExtendedFileProperties() {
            var part = _wordprocessingDocument.ExtendedFilePropertiesPart;
            if (part == null) {
                part = _wordprocessingDocument.AddExtendedFilePropertiesPart();
            }

            if (part.Properties == null) {
                part.Properties = new Properties();
            }
        }
    }
}
