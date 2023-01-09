using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordField {
        public static WordParagraph AddField(WordParagraph paragraph, WordFieldType wordFieldType, WordFieldFormat? wordFieldFormat = null, bool advanced = false, List<String> parameters = null) {
            if (advanced) {
                var runStart = AddFieldStart();
                var runField = AddAdvancedField(wordFieldType: wordFieldType, parameters: parameters);
                var runSeparator = AddFieldSeparator();
                var runText = AddFieldText(wordFieldType.ToString());
                var runEnd = AddFieldEnd();

                paragraph._paragraph.Append(runStart);
                paragraph._paragraph.Append(runField);
                paragraph._paragraph.Append(runSeparator);
                paragraph._paragraph.Append(runText);
                paragraph._paragraph.Append(runEnd);
                paragraph._runs = new List<Run>() { runStart, runField, runSeparator, runText, runEnd };
            } else {
                var simpleField = AddSimpleField(wordFieldType, wordFieldFormat, parameters: parameters);
                paragraph._paragraph.Append(simpleField);
                paragraph._simpleField = simpleField;
            }

            paragraph.Field.UpdateField = true;
            return paragraph;
        }

        public void Remove() {
            if (_runs != null) {
                foreach (var run in _runs) {
                    run.Remove();
                }
            }
            if (this._simpleField != null) {
                this._simpleField.Remove();
            }
        }
    }
}
