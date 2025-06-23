using System.Collections.Generic;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents the FILENAME field code.
    /// </summary>
    public class FileNameField : WordFieldCode {
        /// <summary>
        /// When true, the full path of the file is included (\\p switch).
        /// </summary>
        public bool IncludePath { get; set; }

        internal override WordFieldType FieldType => WordFieldType.FileName;

        internal override List<string> GetParameters() {
            var parameters = new List<string>();
            if (IncludePath) {
                parameters.Add("\\p");
            }
            return parameters;
        }
    }
}
