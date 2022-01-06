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
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Application.Text;
            }
            set {
                CreateExtendedFileProperties();
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Application.Text = value;
            }
        }
        public string ApplicationVersion {
            get {
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.ApplicationVersion.Text;
            }
            set {
                CreateExtendedFileProperties();
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.ApplicationVersion.Text = value;
            }
        }
        public string Paragraphs {
            get {
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Paragraphs.Text;
            }
            set {
                CreateExtendedFileProperties();
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Paragraphs.Text = value;
            }
        }
        public string Characters {
            get {
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Characters.Text;
            }
            set {
                CreateExtendedFileProperties();
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Characters.Text = value;
            }
        }
        public string CharactersWithSpaces {
            get {
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.CharactersWithSpaces.Text;
            }
            set {
                CreateExtendedFileProperties();
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.CharactersWithSpaces.Text = value;
            }
        }
        public string Company {
            get {
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Company.Text;
            }
            set {
                CreateExtendedFileProperties();
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Company.Text = value;
            }
        }
        public DigitalSignature DigitalSignature {
            get {
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.DigitalSignature;
            }
            set {
                CreateExtendedFileProperties();
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.DigitalSignature = value;
            }
        }
        public DocumentSecurity DocumentSecurity {
            get {
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.DocumentSecurity;
            }
            set {
                CreateExtendedFileProperties();
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.DocumentSecurity = value;
            }
        }
        public HeadingPairs HeadingPairs {
            get {
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.HeadingPairs;
            }
            set {
                CreateExtendedFileProperties();
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.HeadingPairs = value;
            }
        }
        public HiddenSlides HiddenSlides {
            get {
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.HiddenSlides;
            }
            set {
                CreateExtendedFileProperties();
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.HiddenSlides = value;
            }
        }
        public HyperlinkBase HyperlinkBase {
            get {
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.HyperlinkBase;
            }
            set {
                CreateExtendedFileProperties();
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.HyperlinkBase = value;
            }
        }
        public HyperlinkList HyperlinkList {
            get {
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.HyperlinkList;
            }
            set {
                CreateExtendedFileProperties();
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.HyperlinkList = value;
            }
        }
        public Lines Lines {
            get {
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Lines;
            }
            set {
                CreateExtendedFileProperties();
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Lines = value;
            }
        }
        public Manager Manager {
            get {
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Manager;
            }
            set {
                CreateExtendedFileProperties();
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Manager = value;
            }
        }
        public HyperlinksChanged HyperlinksChanged {
            get {
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.HyperlinksChanged;
            }
            set {
                CreateExtendedFileProperties();
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.HyperlinksChanged = value;
            }
        }
        public Notes Notes {
            get {
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Notes;
            }
            set {
                CreateExtendedFileProperties();
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Notes = value;
            }
        }
        public MultimediaClips MultimediaClips {
            get {
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.MultimediaClips;
            }
            set {
                CreateExtendedFileProperties();
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.MultimediaClips = value;
            }
        }
        public TotalTime TotalTime {
            get {
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.TotalTime;
            }
            set {
                CreateExtendedFileProperties();
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.TotalTime = value;
            }
        }
        public ScaleCrop ScaleCrop {
            get {
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.ScaleCrop;
            }
            set {
                CreateExtendedFileProperties();
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.ScaleCrop = value;
            }
        }
        public PresentationFormat PresentationFormat {
            get {
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.PresentationFormat;
            }
            set {
                CreateExtendedFileProperties();
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.PresentationFormat = value;
            }
        }
        public Template Template {
            get {
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Template;
            }
            set {
                CreateExtendedFileProperties();
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Template = value;
            }
        }

        public SharedDocument SharedDocument {
            get {
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.SharedDocument;
            }
            set {
                CreateExtendedFileProperties();
                _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.SharedDocument = value;
            }
        }

        public Words Words {
            get {
                return _wordprocessingDocument.ExtendedFilePropertiesPart.Properties.Words;
            }
            set {
                CreateExtendedFileProperties();
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
            }
        }
    }
}
