using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordField {
        private static SimpleField AddSimpleField(WordFieldType wordFieldType, WordFieldFormat? wordFieldFormat = null, List<String> parameters = null) {
            SimpleField simpleField1 = new SimpleField() { Instruction = GenerateField(wordFieldType, wordFieldFormat, parameters) };

            Run run1 = new Run();

            RunProperties runProperties = new RunProperties();
            NoProof noProof = new NoProof();

            runProperties.Append(noProof);
            Text text = new Text {
                Text = "[Document " + wordFieldType + "]"
            };

            run1.Append(runProperties);
            run1.Append(text);

            simpleField1.Append(run1);
            return simpleField1;
        }


        private static Run AddAdvancedField(WordFieldType wordFieldType, WordFieldFormat? wordFieldFormat = null, List<String> parameters = null) {
            Run run = new Run();

            RunProperties runProperties = new RunProperties();
            runProperties.Append(new NoProof());

            FieldCode fieldCode1 = new FieldCode {
                Space = SpaceProcessingModeValues.Preserve,
                Text = GenerateField(wordFieldType, wordFieldFormat, parameters)
            };

            run.Append(runProperties);
            run.Append(fieldCode1);
            return run;
        }

        private static string GenerateField(WordFieldType wordFieldType, WordFieldFormat? wordFieldFormat = null, List<String> parameters = null) {
            var fieldType = " " + wordFieldType.ToString().ToUpper() + " ";
            var fieldFormat = "";
            if (wordFieldFormat != null) {
                fieldFormat = @"\* " + wordFieldFormat + " ";
            }

            var switchesList = " ";
            if (parameters != null) {
                switchesList += parameters.Select(s => s.Trim()).Aggregate((s1, s2) => s1 + ' ' + s2);
            }

            return fieldType + switchesList + fieldFormat + @"\* MERGEFORMAT ";
        }

        private static Run AddFieldSeparator() {
            Run run1 = new Run();

            RunProperties runProperties1 = new RunProperties();
            NoProof noProof1 = new NoProof();

            runProperties1.Append(noProof1);
            FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

            run1.Append(runProperties1);
            run1.Append(fieldChar1);
            return run1;
        }

        private static Run AddFieldStart() {
            Run run1 = new Run();

            RunProperties runProperties1 = new RunProperties();
            NoProof noProof1 = new NoProof();

            runProperties1.Append(noProof1);
            FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

            run1.Append(runProperties1);
            run1.Append(fieldChar1);
            return run1;
        }

        private static Run AddFieldText(string wordType) {
            Run run1 = new Run();

            RunProperties runProperties1 = new RunProperties();
            NoProof noProof1 = new NoProof();

            runProperties1.Append(noProof1);
            Text text1 = new Text();
            text1.Text = "[Document " + wordType + "]";

            run1.Append(runProperties1);
            run1.Append(text1);
            return run1;
        }

        private static Run AddFieldEnd() {
            Run run1 = new Run();

            RunProperties runProperties1 = new RunProperties();
            NoProof noProof1 = new NoProof();

            runProperties1.Append(noProof1);
            FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.End };

            run1.Append(runProperties1);
            run1.Append(fieldChar1);
            return run1;
        }

    }
}