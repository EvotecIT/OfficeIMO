using System;
using System.Diagnostics;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        // internal header properties for easy usage
        internal Header _headerFirst;
        internal Header _headerEven;
        internal Header _headerDefault;
        // internal footer properties for easy usage
        internal Footer _footerFirst;
        internal Footer _footerDefault;
        internal Footer _footerEven;

        //public readonly WordFooters Footer = new WordFooters();
        //public readonly WordHeaders Header = new WordHeaders();

        /// <summary>
        /// Gets or sets the Header.
        /// </summary>
        public WordHeaders Header {
            get {
                if (this.Sections.Count > 1) {
                    Debug.WriteLine("This document contains more than 1 section. Consider using Sections[wantedSection].Headers.");
                }
                return this.Sections[0].Header;
            }
        }
        /// <summary>
        /// Gets or sets the Footer.
        /// </summary>
        public WordFooters Footer {
            get {
                if (this.Sections.Count > 1) {
                    Debug.WriteLine("This document contains more than 1 section. Consider using Sections[wantedSection].Headers.");
                }
                return this.Sections[0].Footer;
            }
        }

        /// <summary>
        /// Gets or sets the DifferentFirstPage.
        /// </summary>
        public bool DifferentFirstPage {
            get {
                if (this.Sections.Count > 1) {
                    Debug.WriteLine("This document contains more than 1 section. Consider using Sections[wantedSection].Headers.");
                }
                return this.Sections[0].DifferentFirstPage;
            }
            set {
                this.Sections[0].DifferentFirstPage = value;
            }

            //get {
            //    var sectionProperties = _document.Body.ChildElements.OfType<SectionProperties>().FirstOrDefault();
            //    if (sectionProperties != null) {
            //        var titlePage = sectionProperties.ChildElements.OfType<TitlePage>().FirstOrDefault();
            //        if (titlePage != null) {
            //            return true;
            //        }
            //    }
            //    return false;
            //}
            //set {
            //    var sectionProperties = _document.Body.ChildElements.OfType<SectionProperties>().FirstOrDefault();
            //    if (sectionProperties == null) {
            //        if (value == false) {
            //            // section properties doesn't exists, so we don't do anything
            //            return;
            //        } else {
            //            _document.Body.Append(
            //               //WordHeadersAndFooters.AddSectionProperties()
            //            );
            //        }
            //    }

            //    sectionProperties = _document.Body.ChildElements.OfType<SectionProperties>().First();
            //    var titlePage = sectionProperties.ChildElements.OfType<TitlePage>().FirstOrDefault();
            //    if (value == false) {
            //        if (titlePage == null) {
            //            return;
            //        } else {
            //            titlePage.Remove();
            //        }
            //    } else {
            //        sectionProperties.Append(new TitlePage());
            //    }

            //}

        }
        /// <summary>
        /// Gets or sets the DifferentOddAndEvenPages.
        /// </summary>
        public bool DifferentOddAndEvenPages {
            get {
                if (this.Sections.Count > 1) {
                    Debug.WriteLine("This document contains more than 1 section. Consider using Sections[wantedSection].Headers.");
                }
                return this.Sections[0].DifferentOddAndEvenPages;
            }
            set {
                this.Sections[0].DifferentOddAndEvenPages = value;
            }
            //get {
            //    var settings = _wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings.ChildElements.OfType<EvenAndOddHeaders>().FirstOrDefault();
            //    if (settings != null) {
            //        return true;
            //    } else {
            //        return false;
            //    }
            //}
            //set {
            //    var settings = _wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings.ChildElements.OfType<EvenAndOddHeaders>().FirstOrDefault();
            //    if (value == false) {

            //    } else {
            //        if (settings == null) {
            //            _wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings.Append(new EvenAndOddHeaders());
            //        } else {
            //            // noting to do, already enabled
            //        }
            //    }
            //}
        }

    }
}
