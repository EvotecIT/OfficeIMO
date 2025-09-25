using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides strongly typed access to the extended application properties
    /// stored in the underlying <see cref="WordprocessingDocument"/>.
    /// </summary>
    public class ApplicationProperties {
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
        /// Gets or sets the digital signature information for the document.
        /// </summary>
        public DigitalSignature? DigitalSignature {
            get => _wordprocessingDocument.ExtendedFilePropertiesPart?.Properties?.DigitalSignature;
            set {
                var properties = GetProperties();
                if (properties != null) {
                    properties.DigitalSignature = value;
                }
            }
        }
        /// <summary>
        /// Gets or sets the document security information.
        /// </summary>
        public DocumentSecurity? DocumentSecurity {
            get => _wordprocessingDocument.ExtendedFilePropertiesPart?.Properties?.DocumentSecurity;
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }
                properties.DocumentSecurity = value;
            }
        }
        /// <summary>
        /// Gets or sets the heading pairs associated with the document.
        /// </summary>
        public HeadingPairs? HeadingPairs {
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
        public HiddenSlides? HiddenSlides {
            get => _wordprocessingDocument.ExtendedFilePropertiesPart?.Properties?.HiddenSlides;
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }
                properties.HiddenSlides = value;
            }
        }
        /// <summary>
        /// Gets or sets the base address used for resolving hyperlinks.
        /// </summary>
        public HyperlinkBase? HyperlinkBase {
            get => _wordprocessingDocument.ExtendedFilePropertiesPart?.Properties?.HyperlinkBase;
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }
                properties.HyperlinkBase = value;
            }
        }
        /// <summary>
        /// Gets or sets the list of hyperlinks in the document.
        /// </summary>
        public HyperlinkList? HyperlinkList {
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
        public Lines? Lines {
            get => _wordprocessingDocument.ExtendedFilePropertiesPart?.Properties?.Lines;
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }
                properties.Lines = value;
            }
        }
        /// <summary>
        /// Gets or sets the manager associated with the document.
        /// </summary>
        public Manager? Manager {
            get => _wordprocessingDocument.ExtendedFilePropertiesPart?.Properties?.Manager;
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }
                properties.Manager = value;
            }
        }
        /// <summary>
        /// Gets or sets a value indicating whether hyperlinks have changed.
        /// </summary>
        public HyperlinksChanged? HyperlinksChanged {
            get => _wordprocessingDocument.ExtendedFilePropertiesPart?.Properties?.HyperlinksChanged;
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }
                properties.HyperlinksChanged = value;
            }
        }
        /// <summary>
        /// Gets or sets the notes information for the document.
        /// </summary>
        public Notes? Notes {
            get => _wordprocessingDocument.ExtendedFilePropertiesPart?.Properties?.Notes;
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }
                properties.Notes = value;
            }
        }
        /// <summary>
        /// Gets or sets the multimedia clips associated with the document.
        /// </summary>
        public MultimediaClips? MultimediaClips {
            get => _wordprocessingDocument.ExtendedFilePropertiesPart?.Properties?.MultimediaClips;
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }
                properties.MultimediaClips = value;
            }
        }
        /// <summary>
        /// Gets or sets the total editing time for the document.
        /// </summary>
        public TotalTime? TotalTime {
            get => _wordprocessingDocument.ExtendedFilePropertiesPart?.Properties?.TotalTime;
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }
                properties.TotalTime = value;
            }
        }
        /// <summary>
        /// Gets or sets the scale crop information for the document.
        /// </summary>
        public ScaleCrop? ScaleCrop {
            get => _wordprocessingDocument.ExtendedFilePropertiesPart?.Properties?.ScaleCrop;
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }
                properties.ScaleCrop = value;
            }
        }
        /// <summary>
        /// Gets or sets the presentation format used by the document.
        /// </summary>
        public PresentationFormat? PresentationFormat {
            get => _wordprocessingDocument.ExtendedFilePropertiesPart?.Properties?.PresentationFormat;
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }
                properties.PresentationFormat = value;
            }
        }
        /// <summary>
        /// Gets or sets the template from which the document was created.
        /// </summary>
        public Template? Template {
            get => _wordprocessingDocument.ExtendedFilePropertiesPart?.Properties?.Template;
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }
                properties.Template = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the document is shared.
        /// </summary>
        public SharedDocument? SharedDocument {
            get => _wordprocessingDocument.ExtendedFilePropertiesPart?.Properties?.SharedDocument;
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }
                properties.SharedDocument = value;
            }
        }

        /// <summary>
        /// Gets or sets the total number of words in the document.
        /// </summary>
        public Words? Words {
            get => _wordprocessingDocument.ExtendedFilePropertiesPart?.Properties?.Words;
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }
                properties.Words = value;
            }
        }

        /// <summary>
        /// Initializes a new instance bound to the specified document.
        /// </summary>
        /// <param name="document">Parent document.</param>
        public ApplicationProperties(WordDocument document) {
            _document = document;
            _wordprocessingDocument = document._wordprocessingDocument;
            _document.ApplicationProperties = this;
        }

        private Properties? GetProperties() {
            CreateExtendedFileProperties();
            return _wordprocessingDocument.ExtendedFilePropertiesPart?.Properties;
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
