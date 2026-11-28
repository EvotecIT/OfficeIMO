using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word {
    // Provides public methods for field manipulation.
    public partial class WordField {
        /// <summary>
        /// Inserts a field into the specified paragraph.
        /// </summary>
        /// <param name="paragraph">Paragraph to add the field to.</param>
        /// <param name="wordFieldType">Type of the field.</param>
        /// <param name="wordFieldFormat">Optional field format.</param>
        /// <param name="customFormat">Custom format string for date or time fields.</param>
        /// <param name="advanced">Whether to use advanced field representation.</param>
        /// <param name="parameters">Additional switch parameters.</param>
        /// <returns>The created field.</returns>
        public static WordField AddField(WordParagraph paragraph, WordFieldType wordFieldType, WordFieldFormat? wordFieldFormat = null, string? customFormat = null, bool advanced = false, List<string>? parameters = null) {
            OfficeXmlNode parent = paragraph.EditParagraphForField();
            string instruction = GenerateField(wordFieldType, wordFieldFormat, customFormat, parameters);
            string text = "[Document " + wordFieldType + "]";
            WordField field;
            if (advanced) {
                var runs = new List<OfficeXmlNode>(5) {
                    AppendMarkerRun(parent, "begin"),
                    AppendTextRun(parent, "instrText", instruction, preserveSpace: true),
                    AppendMarkerRun(parent, "separate"),
                    AppendTextRun(parent, "t", text, preserveSpace: false),
                    AppendMarkerRun(parent, "end")
                };
                field = new WordField(paragraph._document, parent, default, runs);
            } else {
                OfficeXmlNode simpleField = parent.CreateElement(W, "fldSimple");
                simpleField.SetAttribute(W, "instr", instruction);
                AppendTextRun(simpleField, "t", text, preserveSpace: false);
                field = new WordField(paragraph._document, parent, parent.Append(simpleField));
            }

            if (parent.Tree != null && paragraph._document._ownedModel is { } model && ReferenceEquals(model.Document, parent.Tree))
                model.NoteBodyFieldAdded();
            paragraph.RememberAuthoredField(field);
            field._authoringParagraph = paragraph;
            field.UpdateField = true;
            return field;
        }

        /// <summary>
        /// Inserts a field represented by a <see cref="WordFieldCode"/> into the specified paragraph.
        /// </summary>
        /// <param name="paragraph">Paragraph to add the field to.</param>
        /// <param name="fieldCode">Field code instance describing instructions and switches.</param>
        /// <param name="wordFieldFormat">Optional field format.</param>
        /// <param name="customFormat">Custom format string for date or time fields.</param>
        /// <param name="advanced">Whether to use advanced field representation.</param>
        /// <returns>The created field.</returns>
        public static WordField AddField(WordParagraph paragraph, WordFieldCode fieldCode, WordFieldFormat? wordFieldFormat = null, string? customFormat = null, bool advanced = false) {
            ArgumentNullException.ThrowIfNull(fieldCode);

            var parameters = fieldCode.GetParameters();
            return AddField(paragraph, fieldCode.FieldType, wordFieldFormat, customFormat, advanced, parameters);
        }

        /// <summary>
        /// Inserts a field constructed using <see cref="WordFieldBuilder"/>.
        /// </summary>
        /// <param name="paragraph">Paragraph to add the field to.</param>
        /// <param name="builder">Field builder instance.</param>
        /// <param name="advanced">Whether to use advanced field representation.</param>
        /// <returns>The created field.</returns>
        public static WordField AddField(WordParagraph paragraph, WordFieldBuilder builder, bool advanced = false) {
            ArgumentNullException.ThrowIfNull(builder);

            var parameters = builder.GetParameters();
            return AddField(paragraph, builder.FieldType, builder.Format, builder.CustomFormat, advanced, parameters);
        }

        /// <summary>
        /// Deletes all runs and simple field elements associated with this field.
        /// </summary>
        public void Remove() {
            ResolveNodes()?.MarkChanged();
            foreach (OfficeXmlNode run in _runs) run.Remove();
            if (_simpleField.Exists) _simpleField.Remove();
        }
    }
}
