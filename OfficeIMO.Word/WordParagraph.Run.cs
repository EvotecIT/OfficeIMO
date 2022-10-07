using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordParagraph {
        public bool IsEmpty {
            get {
                if (_run == null) {
                    return true;
                }

                return false;
            }
        }
        public bool IsPageBreak {
            get {
                if (this.PageBreak != null) {
                    return true;
                }

                return false;
            }
        }

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
