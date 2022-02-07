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

            //get {
            //    if (_run != null) {
            //        var brake = _run.ChildElements.OfType<Break>().FirstOrDefault();
            //        if (brake != null && brake.Type.Value == BreakValues.Page) {
            //            return true;
            //        } else {
            //            return false;
            //        }
            //    }

            //    return false;
            //}
        }
    }
}
