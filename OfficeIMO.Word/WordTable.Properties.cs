using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Word {
    public partial class WordTable {
        public enum AutoFitType {
            ToContent,
            ToWindow,
            Fixed
        }

        public AutoFitType? AutoFit {
            get {
                return AutoFitType.ToContent;
            }
        }
    }
}
