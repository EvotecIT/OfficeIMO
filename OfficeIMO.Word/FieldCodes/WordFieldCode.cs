using System;
using System.Collections.Generic;

namespace OfficeIMO.Word {
    public abstract class WordFieldCode {
        internal abstract WordFieldType FieldType { get; }

        internal abstract List<string> GetParameters();
    }
}
