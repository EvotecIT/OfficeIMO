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
                    Debug.WriteLine("This document contains more than 1 section. Consider using Sections[wantedSection].Headers.");
                }

                return this.Sections[0].PageOrientation;
            }
            set {
                if (this.Sections.Count > 1) {
                    Debug.WriteLine("This document contains more than 1 section. Consider using Sections[wantedSection].Headers.");
                }

                this.Sections[0].PageOrientation = value;
            }
        }
    }
}