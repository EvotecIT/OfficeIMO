using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {

    /// <summary>
    /// List of supported FieldCodes for Word. For the correlating switches, please have a look at the MS docs:<br/>
    /// <see>https://support.microsoft.com/en-us/office/list-of-field-codes-in-word-1ad6d91a-55a7-4a8d-b535-cf7888659a51 </see>
    /// </summary>
    public enum WordFieldType {
        AddressBlock,
        Advance,
        Ask,
        Author,
        AutoNum,
        AutoNumLgl,
        AutoNumOut,
        AutoText,
        AutoTextList,
        Bibliography,
        Citation,
        Comments,
        Compare,
        CreateDate,
        Database,
        Date,
        DocProperty,
        DocVariable,
        Embed,
        Eq,
        FileName,
        FileSize,
        GoToButton,
        GreetingLine,
        HyperlinkIf,
        IncludePicture,
        IncludeText,
        Index,
        Info,
        Keywords,
        LastSavedBy,
        Link,
        ListNum,
        MacroButton,
        MergeField,
        MergeRec,
        MergeSeq,
        Next,
        NextIf,
        NoteRef,
        NumChars,
        NumPages,
        NumWords,
        Page,
        PageRef,
        Print,
        PrintDate,
        Private,
        Quote,
        RD,
        Ref,
        RevNum,
        SaveDate,
        Section,
        SectionPages,
        Seq,
        Set,
        SkipIf,
        StyleRef,
        Subject,
        Symbol,
        TA,
        TC,
        Template,
        Time,
        Title,
        TOA,
        TOC,
        UserAddress,
        UserInitials,
        UserName,
        XE
    }

    public enum WordFieldFormat {
        Lower,
        Upper,
        FirstCap,
        Caps
    }

    public partial class WordField {
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

        /// <summary>
        /// Format switches, which are identified by <code>\*</code>
        /// </summary>
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

        public List<String> FieldSwitches{
            get
            {
                string pattern = "\\[A-Za-z0-9] +\"?\"";
                Regex rgx = new Regex(pattern);
                var fieldParts = rgx.Split(Field);

                return new List<String>();
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

        //public WordField(WordDocument document, Paragraph paragraph, List<Run> runs) {
        //    this._document = document;
        //    this._paragraph = paragraph;
        //    this._runs = runs;
        //}

        //public WordField(WordDocument document, Paragraph paragraph, SimpleField simpleField) {
        //    this._document = document;
        //    this._paragraph = paragraph;
        //    this._simpleField = simpleField;
        //    this._runs.Add(simpleField.GetFirstChild<Run>());
        //}

        internal WordField(WordDocument document, Paragraph paragraph, SimpleField simpleField, List<Run> runs) {
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
    }
}
