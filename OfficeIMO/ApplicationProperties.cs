using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO {
    public class ApplicationProperties {
        public WordprocessingDocument _wordprocessingDocument = null;
        public WordDocument _document = null;

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
                CreateExtendedFileProperties();
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Application == null) {
                    _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Application = new Application();
                }

                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Application.Text = value;
            }
        }
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
                CreateExtendedFileProperties();
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.ApplicationVersion == null) {
                    _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.ApplicationVersion = new ApplicationVersion();
                }
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.ApplicationVersion.Text = value;
            }
        }
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
                CreateExtendedFileProperties();
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Paragraphs == null) {
                    _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Paragraphs = new Paragraphs();
                }
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Paragraphs.Text = value;
            }
        }
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
                CreateExtendedFileProperties();
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Paragraphs == null) {
                    _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Paragraphs = new Paragraphs();
                }
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Characters.Text = value;
            }
        }
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
                CreateExtendedFileProperties();
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.CharactersWithSpaces == null) {
                    _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.CharactersWithSpaces = new CharactersWithSpaces();
                }
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.CharactersWithSpaces.Text = value;
            }
        }
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
                CreateExtendedFileProperties();
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Company == null) {
                    _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Company = new Company();
                }
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Company.Text = value;
            }
        }
        public DigitalSignature DigitalSignature {
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
                CreateExtendedFileProperties();
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.DigitalSignature == null) {
                    _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.DigitalSignature = new DigitalSignature();
                }
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.DigitalSignature = value;
            }
        }
        public DocumentSecurity DocumentSecurity {
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
                CreateExtendedFileProperties();
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.DocumentSecurity == null) {
                    _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.DocumentSecurity = new DocumentSecurity();
                }
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.DocumentSecurity = value;
            }
        }
        public HeadingPairs HeadingPairs {
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
                CreateExtendedFileProperties();
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.HeadingPairs == null) {
                    _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.HeadingPairs = new HeadingPairs();
                }
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.HeadingPairs = value;
            }
        }
        public HiddenSlides HiddenSlides {
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
                CreateExtendedFileProperties();
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.HiddenSlides == null) {
                    _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.HiddenSlides = new HiddenSlides();
                }
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.HiddenSlides = value;
            }
        }
        public HyperlinkBase HyperlinkBase {
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
                CreateExtendedFileProperties();
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.HyperlinkBase == null) {
                    _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.HyperlinkBase = new HyperlinkBase();
                }
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.HyperlinkBase = value;
            }
        }
        public HyperlinkList HyperlinkList {
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
                CreateExtendedFileProperties();
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.HyperlinkList == null) {
                    _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.HyperlinkList = new HyperlinkList();
                }
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.HyperlinkList = value;
            }
        }
        public Lines Lines {
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
                CreateExtendedFileProperties();
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Lines == null) {
                    _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Lines = new Lines();
                }
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Lines = value;
            }
        }
        public Manager Manager {
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
                CreateExtendedFileProperties();
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Manager == null) {
                    _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Manager = new Manager();
                }
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Manager = value;
            }
        }
        public HyperlinksChanged HyperlinksChanged {
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
                CreateExtendedFileProperties();
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.HyperlinksChanged == null) {
                    _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.HyperlinksChanged = new HyperlinksChanged();
                }
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.HyperlinksChanged = value;
            }
        }
        public Notes Notes {
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
                CreateExtendedFileProperties();
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Notes == null) {
                    _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Notes = new Notes();
                }

                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Notes = value;
            }
        }
        public MultimediaClips MultimediaClips {
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
                CreateExtendedFileProperties();
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.MultimediaClips == null) {
                    _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.MultimediaClips = new MultimediaClips();
                }

                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.MultimediaClips = value;
            }
        }
        public TotalTime TotalTime {
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
                CreateExtendedFileProperties();
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.TotalTime == null) {
                    _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.TotalTime = new TotalTime();
                }
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.TotalTime = value;
            }
        }
        public ScaleCrop ScaleCrop {
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
                CreateExtendedFileProperties();
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.ScaleCrop == null) {
                    _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.ScaleCrop = new ScaleCrop();
                }
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.ScaleCrop = value;
            }
        }
        public PresentationFormat PresentationFormat {
            get {
                if (_wordprocessingDocument.ExtendedFilePropertiesPart == null || _wordprocessingDocument.ExtendedFilePropertiesPart.Properties == null) {
                    return null;
                }
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.PresentationFormat;
            }
            set {
                CreateExtendedFileProperties();
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.PresentationFormat == null) {
                    _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.PresentationFormat = new PresentationFormat();
                }
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.PresentationFormat = value;
            }
        }
        public Template Template {
            get {
                if (_wordprocessingDocument.ExtendedFilePropertiesPart == null || _wordprocessingDocument.ExtendedFilePropertiesPart.Properties == null) {
                    return null;
                }
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Template;
            }
            set {
                CreateExtendedFileProperties();
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Template == null) {
                    _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Template = new Template();
                }
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Template = value;
            }
        }

        public SharedDocument SharedDocument {
            get {
                if (_wordprocessingDocument.ExtendedFilePropertiesPart == null || _wordprocessingDocument.ExtendedFilePropertiesPart.Properties == null) {
                    return null;
                }
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.SharedDocument;
            }
            set {
                CreateExtendedFileProperties();
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.SharedDocument == null) {
                    _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.SharedDocument = new SharedDocument();
                }
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.SharedDocument = value;
            }
        }

        public Words Words {
            get {
                if (_wordprocessingDocument.ExtendedFilePropertiesPart == null || _wordprocessingDocument.ExtendedFilePropertiesPart.Properties == null) {
                    return null;
                }
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Words;
            }
            set {
                CreateExtendedFileProperties();
                if (_wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Words == null) {
                    _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Words = new Words();
                }
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Words = value;
            }
        }

        public ApplicationProperties(WordDocument document) {
            _document = document;
            _wordprocessingDocument = document._wordprocessingDocument;
            _document.ApplicationProperties = this;
        }

        private void CreateExtendedFileProperties() {
            if (_wordprocessingDocument.ExtendedFilePropertiesPart == null) {
                _wordprocessingDocument.AddExtendedFilePropertiesPart();
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties = new Properties();
            }
        }
    }
}
