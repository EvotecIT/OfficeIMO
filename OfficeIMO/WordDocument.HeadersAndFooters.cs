using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO {
    public partial class WordDocument {
        // internal header properties for easy usage
        internal Header _headerFirst;
        internal Header _headerEven;
        internal Header _headerOdd;
        // internal footer properties for easy usage
        internal Footer _footerFirst;
        internal Footer _footerOdd;
        internal Footer _footerEven;

        public readonly WordFooters Footer = new WordFooters();
        public readonly WordHeaders Header = new WordHeaders();

        public bool DifferentFirstPage {
            get {
                var sectionProperties = _document.Body.ChildElements.OfType<SectionProperties>().FirstOrDefault();
                if (sectionProperties != null) {
                    var titlePage = sectionProperties.ChildElements.OfType<TitlePage>().FirstOrDefault();
                    if (titlePage != null) {
                        return true;
                    }
                }
                return false;
            }
            set {
                var sectionProperties = _document.Body.ChildElements.OfType<SectionProperties>().FirstOrDefault();
                if (sectionProperties == null) {
                    if (value == false) {
                        // section properties doesn't exists, so we don't do anything
                        return;
                    } else {
                        _document.Body.Append(
                           //WordHeadersAndFooters.AddSectionProperties()
                        );
                    }
                }

                sectionProperties = _document.Body.ChildElements.OfType<SectionProperties>().First();
                var titlePage = sectionProperties.ChildElements.OfType<TitlePage>().FirstOrDefault();
                if (value == false) {
                    if (titlePage == null) {
                        return;
                    } else {
                        titlePage.Remove();
                    }
                } else {
                    sectionProperties.Append(new TitlePage());
                }

            }

        }

        public bool DifferentOddAndEvenPages {
            get {
                var settings = _wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings.ChildElements.OfType<EvenAndOddHeaders>().FirstOrDefault();
                if (settings != null) {
                    return true;
                } else {
                    return false;
                }
            }
            set {
                var settings = _wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings.ChildElements.OfType<EvenAndOddHeaders>().FirstOrDefault();
                if (value == false) {

                } else {
                    if (settings == null) {
                        _wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings.Append(new EvenAndOddHeaders());
                    } else {
                        // noting to do, already enabled
                    }
                }
            }
        }

    }
}
