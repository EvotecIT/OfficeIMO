using System;
using System.Collections.Generic;
using System.Text;
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
                if (_wordprocessingDocument.ExtendedFilePropertiesPart == null || _wordprocessingDocument.ExtendedFilePropertiesPart.Properties == null) {
                    return "";
                }
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Application == null) {
                    return "";
                }
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Application.Text;
            }
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }
                if (properties.Application == null) {
                    properties.Application = new Application();
                }

                properties.Application.Text = value;
            }
        }
        /// <summary>
        /// Gets or sets the version of the application that created the document.
        /// </summary>
        public string ApplicationVersion {
            get {
                if (_wordprocessingDocument.ExtendedFilePropertiesPart == null || _wordprocessingDocument.ExtendedFilePropertiesPart.Properties == null) {
                    return "";
                }
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.ApplicationVersion == null) {
                    return "";
                }
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.ApplicationVersion.Text;
            }
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }
                if (properties.ApplicationVersion == null) {
                    properties.ApplicationVersion = new ApplicationVersion();
                }
                properties.ApplicationVersion.Text = value;
            }
        }
        /// <summary>
        /// Gets or sets the total number of paragraphs in the document.
        /// </summary>
        public string Paragraphs {
            get {
                if (_wordprocessingDocument.ExtendedFilePropertiesPart == null || _wordprocessingDocument.ExtendedFilePropertiesPart.Properties == null) {
                    return "";
                }
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Paragraphs == null) {
                    return "";
                }
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Paragraphs.Text;
            }
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }
                if (properties.Paragraphs == null) {
                    properties.Paragraphs = new Paragraphs();
                }
                properties.Paragraphs.Text = value;
            }
        }
        /// <summary>
        /// Gets or sets the total number of pages in the document.
        /// </summary>
        public string Pages {
            get {
                if (_wordprocessingDocument.ExtendedFilePropertiesPart == null || _wordprocessingDocument.ExtendedFilePropertiesPart.Properties == null) {
                    return "";
                }
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Pages == null) {
                    return "";
                }
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Pages.Text;
            }
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }
                if (properties.Pages == null) {
                    properties.Pages = new Pages();
                }
                properties.Pages.Text = value;
            }
        }
        /// <summary>
        /// Gets or sets the character count of the document.
        /// </summary>
        public string Characters {
            get {
                if (_wordprocessingDocument.ExtendedFilePropertiesPart == null || _wordprocessingDocument.ExtendedFilePropertiesPart.Properties == null) {
                    return "";
                }
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Characters == null) {
                    return "";
                }
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Characters.Text;
            }
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }
                if (properties.Characters == null) {
                    properties.Characters = new Characters();
                }
                properties.Characters.Text = value;
            }
        }
        /// <summary>
        /// Gets or sets the character count including spaces.
        /// </summary>
        public string CharactersWithSpaces {
            get {
                if (_wordprocessingDocument.ExtendedFilePropertiesPart == null || _wordprocessingDocument.ExtendedFilePropertiesPart.Properties == null) {
                    return "";
                }
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.CharactersWithSpaces == null) {
                    return "";
                }
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.CharactersWithSpaces.Text;
            }
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }
                if (properties.CharactersWithSpaces == null) {
                    properties.CharactersWithSpaces = new CharactersWithSpaces();
                }
                properties.CharactersWithSpaces.Text = value;
            }
        }
        /// <summary>
        /// Gets or sets the company associated with the document.
        /// </summary>
        public string Company {
            get {
                if (_wordprocessingDocument.ExtendedFilePropertiesPart == null || _wordprocessingDocument.ExtendedFilePropertiesPart.Properties == null) {
                    return "";
                }
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Company == null) {
                    return "";
                }
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Company.Text;
            }
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }
                if (properties.Company == null) {
                    properties.Company = new Company();
                }
                properties.Company.Text = value;
            }
        }
        /// <summary>
        /// Gets or sets the digital signature information for the document.
        /// </summary>
        public DigitalSignature? DigitalSignature {
            get {
                if (_wordprocessingDocument.ExtendedFilePropertiesPart == null || _wordprocessingDocument.ExtendedFilePropertiesPart.Properties == null) {
                    return null;
                }
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.DigitalSignature == null) {
                    return null;
                }
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.DigitalSignature;
            }
            set {
                var properties = GetProperties();
                if (properties == null) {
                    return;
                }
                properties.DigitalSignature = value;
            }
        }
        /// <summary>
        /// Gets or sets the document security information.
        /// </summary>
        public DocumentSecurity? DocumentSecurity {
            get {
                if (_wordprocessingDocument.ExtendedFilePropertiesPart == null || _wordprocessingDocument.ExtendedFilePropertiesPart.Properties == null) {
                    return null;
                }
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.DocumentSecurity == null) {
                    return null;
                }
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.DocumentSecurity;
            }
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
            get {
                if (_wordprocessingDocument.ExtendedFilePropertiesPart == null || _wordprocessingDocument.ExtendedFilePropertiesPart.Properties == null) {
                    return null;
                }
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.HeadingPairs == null) {
                    return null;
                }
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.HeadingPairs;
            }
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
            get {
                if (_wordprocessingDocument.ExtendedFilePropertiesPart == null || _wordprocessingDocument.ExtendedFilePropertiesPart.Properties == null) {
                    return null;
                }
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.HiddenSlides == null) {
                    return null;
                }
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.HiddenSlides;
            }
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
            get {
                if (_wordprocessingDocument.ExtendedFilePropertiesPart == null || _wordprocessingDocument.ExtendedFilePropertiesPart.Properties == null) {
                    return null;
                }
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.HyperlinkBase == null) {
                    return null;
                }
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.HyperlinkBase;
            }
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
            get {
                if (_wordprocessingDocument.ExtendedFilePropertiesPart == null || _wordprocessingDocument.ExtendedFilePropertiesPart.Properties == null) {
                    return null;
                }
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.HyperlinkList == null) {
                    return null;
                }
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.HyperlinkList;
            }
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
            get {
                if (_wordprocessingDocument.ExtendedFilePropertiesPart == null || _wordprocessingDocument.ExtendedFilePropertiesPart.Properties == null) {
                    return null;
                }
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Lines == null) {
                    return null;
                }
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Lines;
            }
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
            get {
                if (_wordprocessingDocument.ExtendedFilePropertiesPart == null || _wordprocessingDocument.ExtendedFilePropertiesPart.Properties == null) {
                    return null;
                }
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Manager == null) {
                    return null;
                }
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Manager;
            }
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
            get {
                if (_wordprocessingDocument.ExtendedFilePropertiesPart == null || _wordprocessingDocument.ExtendedFilePropertiesPart.Properties == null) {
                    return null;
                }
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.HyperlinksChanged == null) {
                    return null;
                }
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.HyperlinksChanged;
            }
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
            get {
                if (_wordprocessingDocument.ExtendedFilePropertiesPart == null || _wordprocessingDocument.ExtendedFilePropertiesPart.Properties == null) {
                    return null;
                }
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Notes == null) {
                    return null;
                }
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Notes;
            }
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
            get {
                if (_wordprocessingDocument.ExtendedFilePropertiesPart == null || _wordprocessingDocument.ExtendedFilePropertiesPart.Properties == null) {
                    return null;
                }
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.MultimediaClips == null) {
                    return null;
                }
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.MultimediaClips;
            }
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
            get {
                if (_wordprocessingDocument.ExtendedFilePropertiesPart == null || _wordprocessingDocument.ExtendedFilePropertiesPart.Properties == null) {
                    return null;
                }
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.TotalTime == null) {
                    return null;
                }
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.TotalTime;
            }
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
            get {
                if (_wordprocessingDocument.ExtendedFilePropertiesPart == null || _wordprocessingDocument.ExtendedFilePropertiesPart.Properties == null) {
                    return null;
                }
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.ScaleCrop == null) {
                    return null;
                }
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.ScaleCrop;
            }
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
            get {
                if (_wordprocessingDocument.ExtendedFilePropertiesPart == null || _wordprocessingDocument.ExtendedFilePropertiesPart.Properties == null) {
                    return null;
                }
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.PresentationFormat;
            }
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
            get {
                if (_wordprocessingDocument.ExtendedFilePropertiesPart == null || _wordprocessingDocument.ExtendedFilePropertiesPart.Properties == null) {
                    return null;
                }
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Template;
            }
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
            get {
                if (_wordprocessingDocument.ExtendedFilePropertiesPart == null || _wordprocessingDocument.ExtendedFilePropertiesPart.Properties == null) {
                    return null;
                }
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.SharedDocument;
            }
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
            get {
                if (_wordprocessingDocument.ExtendedFilePropertiesPart == null || _wordprocessingDocument.ExtendedFilePropertiesPart.Properties == null) {
                    return null;
                }
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Words;
            }
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
