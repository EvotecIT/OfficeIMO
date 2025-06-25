using System;
using System.Collections.Generic;

namespace OfficeIMO.Word {
    /// <summary>
    /// Base class for all field code representations. Derived types expose
    /// specific parameters and the field type they represent.
    /// </summary>
    public abstract class WordFieldCode {
        /// <summary>
        /// Gets the type of field represented by the derived class.
        /// </summary>
        internal abstract WordFieldType FieldType { get; }

        /// <summary>
        /// Retrieves a list of parameter strings used to construct the field
        /// code within the document.
        /// </summary>
        internal abstract List<string> GetParameters();
    }
}
