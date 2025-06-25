using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Exposes run-level helpers for WordParagraph.
    /// </summary>
    public partial class WordParagraph {
        /// <summary>
        /// Gets a value indicating whether the paragraph contains no run element.
        /// </summary>
        public bool IsEmpty {
            get {
                if (_run == null) {
                    return true;
                }

                return false;
            }
        }
        /// <summary>
        /// Gets a value indicating whether the paragraph contains a page break.
        /// </summary>
        public bool IsPageBreak {
            get {
                if (this.PageBreak != null) {
                    return true;
                }

                return false;
            }
        }

        /// <summary>
        /// Gets a value indicating whether the paragraph contains a break element.
        /// </summary>
        public bool IsBreak {
            get {
                if (this.Break != null) {
                    return true;
                }

                return false;
            }
        }
    }
}
