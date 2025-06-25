using System;
using System.Collections.Generic;

namespace OfficeIMO.Word {
    /// <summary>
    /// Base class for specific field code representations.
    /// </summary>
    public abstract class WordFieldCode {
        internal abstract WordFieldType FieldType { get; }

        internal abstract List<string> GetParameters();
    }
}
