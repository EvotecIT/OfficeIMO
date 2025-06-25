using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Exposes section-level properties.
    /// </summary>
    public partial class WordDocument {
        /// <summary>
        /// Gets or sets the orientation of the pages in the first section.
        /// </summary>
        /// <remarks>
        /// When a document contains multiple sections, consider using
        /// <c>Sections[wantedSection].PageOrientation</c> instead.
        /// </remarks>
        public PageOrientationValues PageOrientation {
            get {
                if (this.Sections.Count > 1) {
                    Debug.WriteLine("This document contains more than 1 section. Consider using Sections[wantedSection].PageOrientation.");
                }

                return this.Sections[0].PageOrientation;
            }
            set {
                if (this.Sections.Count > 1) {
                    Debug.WriteLine("This document contains more than 1 section. Consider using Sections[wantedSection].PageOrientation.");
                }

                this.Sections[0].PageOrientation = value;
            }
        }

        /// <summary>
        /// Gets or sets the page borders for the first section.
        /// </summary>
        /// <remarks>
        /// When a document contains multiple sections, use
        /// <c>Sections[wantedSection].Borders</c> to target a specific section.
        /// </remarks>
        public WordBorders Borders {
            get {
                if (this.Sections.Count > 1) {
                    Debug.WriteLine("This document contains more than 1 section. Consider using Sections[wantedSection].Borders.");
                }

                return this.Sections[0].Borders;
            }
            set {
                if (this.Sections.Count > 1) {
                    Debug.WriteLine("This document contains more than 1 section. Consider using Sections[wantedSection].Borders.");
                }

                this.Sections[0].Borders = value;
            }
        }

        /// <summary>
        /// Gets or sets the margins for the first section.
        /// </summary>
        /// <remarks>
        /// Use <c>Sections[wantedSection].Margins</c> when a document has more
        /// than one section.
        /// </remarks>
        public WordMargins Margins {
            get {
                if (this.Sections.Count > 1) {
                    Debug.WriteLine("This document contains more than 1 section. Consider using Sections[wantedSection].Margins.");
                }

                return this.Sections[0].Margins;
            }
            set {
                if (this.Sections.Count > 1) {
                    Debug.WriteLine("This document contains more than 1 section. Consider using Sections[wantedSection].Margins.");
                }

                this.Sections[0].Margins = value;
            }
        }

        /// <summary>
        /// Gets or sets the page size settings for the first section.
        /// </summary>
        /// <remarks>
        /// For multi&#8209;section documents use
        /// <c>Sections[wantedSection].PageSettings</c> instead.
        /// </remarks>
        public WordPageSizes PageSettings {
            get {
                if (this.Sections.Count > 1) {
                    Debug.WriteLine("This document contains more than 1 section. Consider using Sections[wantedSection].PageSettings.");
                }

                return this.Sections[0].PageSettings;
            }
            set {
                if (this.Sections.Count > 1) {
                    Debug.WriteLine("This document contains more than 1 section. Consider using Sections[wantedSection].PageSettings.");
                }

                this.Sections[0].PageSettings = value;
            }
        }

        /// <summary>
        /// Gets or sets the footnote properties for the first section.
        /// </summary>
        /// <remarks>
        /// Use <c>Sections[wantedSection].FootnoteProperties</c> to modify
        /// other sections.
        /// </remarks>
        public FootnoteProperties FootnoteProperties {
            get {
                if (this.Sections.Count > 1) {
                    Debug.WriteLine("This document contains more than 1 section. Consider using Sections[wantedSection].FootnoteProperties.");
                }

                return this.Sections[0].FootnoteProperties;
            }
            set {
                if (this.Sections.Count > 1) {
                    Debug.WriteLine("This document contains more than 1 section. Consider using Sections[wantedSection].FootnoteProperties.");
                }

                this.Sections[0].FootnoteProperties = value;
            }
        }

        /// <summary>
        /// Gets or sets the endnote properties for the first section.
        /// </summary>
        /// <remarks>
        /// Use <c>Sections[wantedSection].EndnoteProperties</c> when working
        /// with multiple sections.
        /// </remarks>
        public EndnoteProperties EndnoteProperties {
            get {
                if (this.Sections.Count > 1) {
                    Debug.WriteLine("This document contains more than 1 section. Consider using Sections[wantedSection].EndnoteProperties.");
                }

                return this.Sections[0].EndnoteProperties;
            }
            set {
                if (this.Sections.Count > 1) {
                    Debug.WriteLine("This document contains more than 1 section. Consider using Sections[wantedSection].EndnoteProperties.");
                }

                this.Sections[0].EndnoteProperties = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the gutter should appear on
        /// the right for right-to-left pages in the first section.
        /// </summary>
        /// <remarks>
        /// For other sections use <c>Sections[wantedSection].RtlGutter</c>.
        /// </remarks>
        public bool RtlGutter {
            get {
                if (this.Sections.Count > 1) {
                    Debug.WriteLine("This document contains more than 1 section. Consider using Sections[wantedSection].RtlGutter.");
                }

                return this.Sections[0].RtlGutter;
            }
            set {
                if (this.Sections.Count > 1) {
                    Debug.WriteLine("This document contains more than 1 section. Consider using Sections[wantedSection].RtlGutter.");
                }

                this.Sections[0].RtlGutter = value;
            }
        }

        /// <summary>
        /// Adds or updates footnote configuration for the first section.
        /// </summary>
        /// <param name="numberingFormat">Numbering format to apply.</param>
        /// <param name="position">Location of footnotes.</param>
        /// <param name="restartNumbering">Restart numbering option.</param>
        /// <param name="startNumber">Starting number.</param>
        public void AddFootnoteProperties(NumberFormatValues? numberingFormat = null,
            FootnotePositionValues? position = null,
            RestartNumberValues? restartNumbering = null,
            int? startNumber = null) {
            this.Sections[0].AddFootnoteProperties(numberingFormat, position, restartNumbering, startNumber);
        }

        /// <summary>
        /// Adds or updates endnote configuration for the first section.
        /// </summary>
        /// <param name="numberingFormat">Numbering format to apply.</param>
        /// <param name="position">Location of endnotes.</param>
        /// <param name="restartNumbering">Restart numbering option.</param>
        /// <param name="startNumber">Starting number.</param>
        public void AddEndnoteProperties(NumberFormatValues? numberingFormat = null,
            EndnotePositionValues? position = null,
            RestartNumberValues? restartNumbering = null,
            int? startNumber = null) {
            this.Sections[0].AddEndnoteProperties(numberingFormat, position, restartNumbering, startNumber);
        }

        /// <summary>
        /// Gets or sets the page numbering configuration for the first section.
        /// </summary>
        /// <remarks>
        /// When a document has multiple sections, access
        /// <c>Sections[wantedSection].PageNumberType</c> instead.
        /// </remarks>
        public PageNumberType PageNumberType {
            get {
                if (this.Sections.Count > 1) {
                    Debug.WriteLine("This document contains more than 1 section. Consider using Sections[wantedSection].PageNumberType.");
                }

                return this.Sections[0].PageNumberType;
            }
            set {
                if (this.Sections.Count > 1) {
                    Debug.WriteLine("This document contains more than 1 section. Consider using Sections[wantedSection].PageNumberType.");
                }

                this.Sections[0].PageNumberType = value;
            }
        }

        /// <summary>
        /// Adds or updates page numbering for the first section.
        /// </summary>
        /// <param name="startNumber">Starting page number.</param>
        /// <param name="format">Number format.</param>
        public void AddPageNumbering(int? startNumber = null, NumberFormatValues? format = null) {
            this.Sections[0].AddPageNumbering(startNumber, format);
        }
    }
}
