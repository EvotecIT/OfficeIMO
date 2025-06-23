using System.Collections.Generic;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents the AUTHOR field code.
    /// </summary>
    public class AuthorField : WordFieldCode {
        /// <summary>
        /// Optional author name to insert instead of the document property.
        /// </summary>
        public string Author { get; set; }

        internal override WordFieldType FieldType => WordFieldType.Author;

        internal override List<string> GetParameters() {
            var parameters = new List<string>();
            if (!string.IsNullOrWhiteSpace(Author)) {
                parameters.Add($"\"{Author}\"");
            }
            return parameters;
        }
    }
}
