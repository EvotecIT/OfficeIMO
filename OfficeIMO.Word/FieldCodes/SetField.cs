namespace OfficeIMO.Word {
    /// <summary>
    /// Represents the SET field code.
    /// </summary>
    public class SetField : WordFieldCode {
        /// <summary>
        /// Name of the bookmark or variable to set.
        /// </summary>
        public string? Bookmark { get; set; }

        /// <summary>
        /// Value to assign to the bookmark.
        /// </summary>
        public string? Value { get; set; }

        internal override WordFieldType FieldType => WordFieldType.Set;

        internal override List<string> GetParameters() {
            var parameters = new List<string>();
            if (!string.IsNullOrWhiteSpace(Bookmark)) {
                parameters.Add(Bookmark!);
            }
            if (!string.IsNullOrWhiteSpace(Value)) {
                parameters.Add($"\"{Value!}\"");
            }
            return parameters;
        }
    }
}
