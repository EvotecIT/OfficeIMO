using System.Collections.Generic;
using DocumentFormat.OpenXml.Bibliography;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents the CITATION field code.
    /// </summary>
    public class CitationField : WordFieldCode {
        /// <summary>
        /// Tag of the source to cite.
        /// </summary>
        public string SourceTag { get; set; }

        internal override WordFieldType FieldType => WordFieldType.Citation;

        internal override List<string> GetParameters() {
            var parameters = new List<string>();
            if (!string.IsNullOrWhiteSpace(SourceTag)) {
                parameters.Add(SourceTag);
            }
            return parameters;
        }
    }
}
