using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides basic mail merge capabilities by replacing <c>MERGEFIELD</c> fields with supplied values.
    /// </summary>
    public static class WordMailMerge {
        /// <summary>
        /// Replaces all MERGEFIELD fields in the given document with provided values.
        /// </summary>
        /// <param name="document">Document to update.</param>
        /// <param name="values">Dictionary with field names and values.</param>
        /// <param name="removeFields">Determines whether the field codes are removed after replacement.</param>
        public static void Execute(WordDocument document, IDictionary<string, string> values, bool removeFields = true) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (values == null) throw new ArgumentNullException(nameof(values));

            foreach (var field in document.Fields.Where(f => f.FieldType == WordFieldType.MergeField)) {
                var parser = new WordFieldParser(field.Field);
                if (parser.Instructions.Count == 0) continue;
                var name = parser.Instructions[0].Trim().Trim('"');
                if (values.TryGetValue(name, out var value)) {
                    if (removeFields) {
                        field.ReplaceWithText(value);
                    } else {
                        field.Text = value;
                    }
                }
            }
        }
    }
}
