using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {

    /// <summary>
    /// Enumerates field codes available in Word documents. For
    /// related format switches see the
    /// <see href="https://support.microsoft.com/en-us/office/list-of-field-codes-in-word-1ad6d91a-55a7-4a8d-b535-cf7888659a51">Microsoft documentation</see>.
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

    /// <summary>
    /// Specifies format switches for Word field codes.
    /// </summary>
    public enum WordFieldFormat {
        Lower,
        Upper,
        FirstCap,
        Caps,
        Mergeformat,
        Roman,
        roman,
        Arabic,
        Alphabetical,
        ALPHABETICAL,
        Ordinal,
        OrdText,
        CardText,
        DollarText,
        Hex,
        CharFormat,
    }

    /// <summary>
    /// Represents a field element within a Word document.
    /// </summary>
    public partial class WordField : WordElement {
        private readonly WordDocument _document;
        private readonly Paragraph _paragraph;
        private readonly List<Run> _runs = new List<Run>();
        private readonly SimpleField _simpleField;

        /// <summary>
        /// Gets the type of the current field.
        /// </summary>
        public WordFieldType? FieldType {
            get {
                var parser = new WordFieldParser(Field);
                return parser.WordFieldType;
            }
        }

        /// <summary>
        /// Gets the format switches applied to the field.
        /// </summary>
        public IReadOnlyList<WordFieldFormat> FieldFormat {
            get {
                var parser = new WordFieldParser(Field);
                return parser.FormatSwitches;
            }
        }

        /// <summary>
        /// Gets the raw switch parameters from the field code.
        /// </summary>
        public List<String> FieldSwitches {
            get {
                var parser = new WordFieldParser(Field);
                return parser.Switches;
            }
        }

        /// <summary>
        /// Gets the instructions portion of the field code.
        /// </summary>
        public List<String> FieldInstructions {
            get {
                var parser = new WordFieldParser(Field);
                return parser.Instructions;
            }
        }

        /// <summary>
        /// Gets the raw field code.
        /// </summary>
        public string Field {
            get {
                if (_simpleField != null) {
                    return _simpleField.Instruction;
                } else {
                    var instruction = "";
                    foreach (var run in _runs) {
                        var fieldCode = run.ChildElements.OfType<FieldCode>().FirstOrDefault();
                        if (fieldCode != null) {
                            instruction += fieldCode.Text;
                        }
                    }
                    return instruction;
                }
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the field is marked dirty.
        /// </summary>
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

        /// <summary>
        /// Gets or sets a value indicating whether the field is locked.
        /// </summary>
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

        /// <summary>
        /// Gets or sets the display text of the field.
        /// </summary>
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

        internal WordField(WordDocument document, Paragraph paragraph, SimpleField simpleField, List<Run> runs) {
            this._document = document;
            this._paragraph = paragraph;
            this._simpleField = simpleField;
            if (simpleField != null) {
                this._runs.Add(simpleField.GetFirstChild<Run>());
            } else {
                this._runs = runs;
            }
        }
    }
}
