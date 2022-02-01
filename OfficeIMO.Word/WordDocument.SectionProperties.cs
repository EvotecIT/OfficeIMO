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
    }
}