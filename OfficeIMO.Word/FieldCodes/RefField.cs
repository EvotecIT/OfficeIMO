using System.Collections.Generic;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents the REF field code.
    /// </summary>
    public class RefField : WordFieldCode {
        /// <summary>
        /// Bookmark to reference.
        /// </summary>
        public string Bookmark { get; set; }

        /// <summary>
        /// When true, inserts a hyperlink to the bookmark (\\h switch).
        /// </summary>
        public bool InsertHyperlink { get; set; }

        internal override WordFieldType FieldType => WordFieldType.Ref;

        internal override List<string> GetParameters() {
            var parameters = new List<string>();
            if (!string.IsNullOrWhiteSpace(Bookmark)) {
                parameters.Add(Bookmark);
            }
            if (InsertHyperlink) {
                parameters.Add("\\h");
            }
            return parameters;
        }
    }
}
