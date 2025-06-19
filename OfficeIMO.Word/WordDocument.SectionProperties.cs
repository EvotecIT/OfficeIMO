using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        /// <summary>
        /// Gets or sets the PageOrientation.
        /// </summary>
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
        /// Gets or sets the Borders.
        /// </summary>
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
        /// Gets or sets the Margins.
        /// </summary>
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
        /// Gets or sets the PageSettings.
        /// </summary>
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
        /// Gets or sets the FootnoteProperties.
        /// </summary>
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
        /// Gets or sets the EndnoteProperties.
        /// </summary>
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
        /// Gets or sets the RtlGutter.
        /// </summary>
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
        /// Applies footnote options to the first section of the document.
        /// </summary>
        public void AddFootnoteProperties(NumberFormatValues? numberingFormat = null,
            FootnotePositionValues? position = null,
            RestartNumberValues? restartNumbering = null,
            int? startNumber = null) {
            this.Sections[0].AddFootnoteProperties(numberingFormat, position, restartNumbering, startNumber);
        }

        /// <summary>
        /// Applies endnote options to the first section of the document.
        /// </summary>
        public void AddEndnoteProperties(NumberFormatValues? numberingFormat = null,
            EndnotePositionValues? position = null,
            RestartNumberValues? restartNumbering = null,
            int? startNumber = null) {
            this.Sections[0].AddEndnoteProperties(numberingFormat, position, restartNumbering, startNumber);
        }
    }
}
