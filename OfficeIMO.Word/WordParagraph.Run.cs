using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordParagraph {
        /// <summary>
        /// Gets or sets the IsEmpty.
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
        /// Gets or sets the IsPageBreak.
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
        /// Gets or sets the IsBreak.
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
