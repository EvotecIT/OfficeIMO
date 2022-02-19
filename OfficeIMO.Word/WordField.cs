using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {

    public enum WordFieldType {
        Comments,
        Page,
        Title,
        Keywords,
        Subject,
        Time,
        Author,
        FileSize,
        FileName
    }

    public enum WordFieldFormat {
        Lower,
        Upper,
        FirstCap,
        Caps
    }

    public class WordField {
        private readonly WordDocument _document;
        private readonly Paragraph _paragraph;
        private readonly List<Run> _runs = new List<Run>();
        private readonly SimpleField _simpleField;

        public WordFieldType? FieldType {
            get {
                var splitField = Field.Split(new string[] { "\\" }, StringSplitOptions.None);
                if (splitField.Length == 3 || splitField.Length == 2) {
                    var fieldType = splitField[0].Replace("*", "").Trim();
                    return ConvertToWordFieldType(fieldType);
                }
                return null;
            }
        }
        public WordFieldFormat? FieldFormat {
            get {
                var splitField = Field.Split(new string[] { "\\" }, StringSplitOptions.None);
                if (splitField.Length == 3) {
                    var format = splitField[1].Replace("*", "").Trim();
                    return ConvertToWordFieldFormat(format);
                }
                return null;
            }
        }

        public string Field {
            get {
                if (_simpleField != null) {
                    return _simpleField.Instruction;
                } else {
                    foreach (var run in _runs) {
                        var fieldCode = run.ChildElements.OfType<FieldCode>().FirstOrDefault();
                        if (fieldCode != null) {
                            return fieldCode.Text;
                        }
                    }
                }

                return null;
            }
        }

        public bool UpdateField {
            get {
                if (_simpleField != null) {
                    if (_simpleField.Dirty != null) {
                        return _simpleField.Dirty;
                    } else {
                        return false;
                    }
                } else {
                    foreach (var run in _runs) {
                        var fieldChar = run.ChildElements.OfType<FieldChar>().FirstOrDefault();
                        if (fieldChar != null) {
                            if (fieldChar.Dirty != null) {
                                return fieldChar.Dirty;
                            } else {
                                return false;
                            }
                        }
                    }
                }
                return false;
            }
            set {
                if (_simpleField != null) {
                    _simpleField.Dirty = value;
                } else {
                    foreach (var run in _runs) {
                        var fieldChar = run.ChildElements.OfType<FieldChar>().FirstOrDefault();
                        if (fieldChar != null) {
                            fieldChar.Dirty = value;
                        }
                    }
                }
            }
        }

        public bool LockField {
            get {
                if (_simpleField != null) {
                    if (_simpleField.FieldLock != null) {
                        return _simpleField.FieldLock;
                    } else {
                        return false;
                    }
                } else {
                    foreach (var run in _runs) {
                        var fieldChar = run.ChildElements.OfType<FieldChar>().FirstOrDefault();
                        if (fieldChar != null) {
                            if (fieldChar.FieldLock != null) {
                                return fieldChar.FieldLock;
                            } else {
                                return false;
                            }
                        }
                    }
                }
                return false;
            }
            set {
                if (_simpleField != null) {
                    _simpleField.FieldLock = value;
                } else {
                    foreach (var run in _runs) {
                        var fieldChar = run.ChildElements.OfType<FieldChar>().FirstOrDefault();
                        if (fieldChar != null) {
                            fieldChar.FieldLock = value;
                        }
                    }
                }
            }
        }

        public string Text {
            get {
                foreach (var run in _runs) {
                    var text = run.ChildElements.OfType<Text>().FirstOrDefault();
                    if (text != null) {
                        return text.Text;
                    }
                }

                return "";
            }
            set {
                foreach (var run in _runs) {
                    var text = run.ChildElements.OfType<Text>().FirstOrDefault();
                    if (text != null) {
                        text.Text = value;
                    }
                }
            }
        }

        public WordField(WordDocument document, Paragraph paragraph, List<Run> runs) {
            this._document = document;
            this._paragraph = paragraph;
            this._runs = runs;
        }

        public WordField(WordDocument document, Paragraph paragraph, SimpleField simpleField) {
            this._document = document;
            this._paragraph = paragraph;
            this._simpleField = simpleField;
            this._runs.Add(simpleField.GetFirstChild<Run>());
        }

        public WordField(WordDocument document, Paragraph paragraph, SimpleField simpleField, List<Run> runs) {
            this._document = document;
            this._paragraph = paragraph;
            this._simpleField = simpleField;
            if (simpleField != null) {
                this._runs.Add(simpleField.GetFirstChild<Run>());
            } else {
                this._runs = runs;
            }
            // this._runs.Add(simpleField.GetFirstChild<Run>());
        }

        public static WordParagraph AddField(WordParagraph paragraph, WordFieldType wordFieldType, WordFieldFormat? wordFieldFormat = null, bool advanced = false) {
            if (advanced) {
                var runStart = AddFieldStart();
                var runField = AddAdvancedField(wordFieldType);
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
                var simpleField = AddSimpleField(wordFieldType, wordFieldFormat);
                paragraph._paragraph.Append(simpleField);
                paragraph._simpleField = simpleField;
            }
            return paragraph;
        }

        private static SimpleField AddSimpleField(WordFieldType wordFieldType, WordFieldFormat? wordFieldFormat = null) {
            SimpleField simpleField1 = new SimpleField() { Instruction = GenerateField(wordFieldType, wordFieldFormat) };

            Run run1 = new Run();

            RunProperties runProperties = new RunProperties();
            NoProof noProof = new NoProof();

            runProperties.Append(noProof);
            Text text = new Text {
                Text = wordFieldType.ToString()
            };

            run1.Append(runProperties);
            run1.Append(text);

            simpleField1.Append(run1);
            return simpleField1;
        }


        private static Run AddAdvancedField(WordFieldType wordFieldType, WordFieldFormat? wordFieldFormat = null) {
            Run run = new Run();

            RunProperties runProperties = new RunProperties();
            runProperties.Append(new NoProof());

            FieldCode fieldCode1 = new FieldCode {
                Space = SpaceProcessingModeValues.Preserve,
                Text = GenerateField(wordFieldType, wordFieldFormat)
            };

            run.Append(runProperties);
            run.Append(fieldCode1);
            return run;
        }

        private static string GenerateField(WordFieldType wordFieldType, WordFieldFormat? wordFieldFormat = null) {
            var fieldType = " " + wordFieldType.ToString().ToUpper() + " ";
            var fieldFormat = "";
            if (wordFieldFormat != null) {
                fieldFormat = "\\* " + wordFieldFormat + " ";
            }
            return fieldType + fieldFormat + "\\* MERGEFORMAT ";
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
            text1.Text = wordType;

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

        private static WordFieldFormat ConvertToWordFieldFormat(string wordFieldFormat) {
            WordFieldFormat myFieldFormat = (WordFieldFormat)Enum.Parse(typeof(WordFieldFormat), wordFieldFormat, true);
            return myFieldFormat;
        }

        private static WordFieldType ConvertToWordFieldType(string wordFieldType) {
            WordFieldType myFieldType = (WordFieldType)Enum.Parse(typeof(WordFieldType), wordFieldType, true);
            return myFieldType;
        }
    }
}
