using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
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

        public void AddFootnoteProperties(NumberFormatValues? numberingFormat = null,
            FootnotePositionValues? position = null,
            RestartNumberValues? restartNumbering = null,
            int? startNumber = null) {
            this.Sections[0].AddFootnoteProperties(numberingFormat, position, restartNumbering, startNumber);
        }

        public void AddEndnoteProperties(NumberFormatValues? numberingFormat = null,
            EndnotePositionValues? position = null,
            RestartNumberValues? restartNumbering = null,
            int? startNumber = null) {
            this.Sections[0].AddEndnoteProperties(numberingFormat, position, restartNumbering, startNumber);
        }

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

        public void AddPageNumbering(int? startNumber = null, NumberFormatValues? format = null) {
            this.Sections[0].AddPageNumbering(startNumber, format);
        }
    }
}
