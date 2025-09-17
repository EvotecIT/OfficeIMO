using System;
using System.Collections.Generic;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents the ASK field code.
    /// </summary>
    public class AskField : WordFieldCode {
        /// <summary>
        /// Bookmark to assign the value to.
        /// </summary>
        public string? Bookmark { get; set; }

        /// <summary>
        /// Prompt displayed to the user.
        /// </summary>
        public string? Prompt { get; set; }

        /// <summary>
        /// Default response for the prompt. Added using the \d switch.
        /// </summary>
        public string? DefaultResponse { get; set; }

        /// <summary>
        /// Indicates whether the prompt should appear only once. Added using the \o switch.
        /// </summary>
        public bool PromptOnce { get; set; }

        internal override WordFieldType FieldType => WordFieldType.Ask;

        internal override List<string> GetParameters() {
            var parameters = new List<string>();
            if (!string.IsNullOrWhiteSpace(Bookmark)) {
                parameters.Add(Bookmark!);
            }
            if (!string.IsNullOrWhiteSpace(Prompt)) {
                parameters.Add($"\"{Prompt!}\"");
            }
            if (!string.IsNullOrWhiteSpace(DefaultResponse)) {
                parameters.Add($"\\d \"{DefaultResponse!}\"");
            }
            if (PromptOnce) {
                parameters.Add("\\o");
            }
            return parameters;
        }
    }
}
