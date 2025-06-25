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
        /// <summary>AddressBlock field code.</summary>
        AddressBlock,
        /// <summary>Advance field code.</summary>
        Advance,
        /// <summary>Ask field code.</summary>
        Ask,
        /// <summary>Author field code.</summary>
        Author,
        /// <summary>AutoNum field code.</summary>
        AutoNum,
        /// <summary>AutoNumLgl field code.</summary>
        AutoNumLgl,
        /// <summary>AutoNumOut field code.</summary>
        AutoNumOut,
        /// <summary>AutoText field code.</summary>
        AutoText,
        /// <summary>AutoTextList field code.</summary>
        AutoTextList,
        /// <summary>Bibliography field code.</summary>
        Bibliography,
        /// <summary>Citation field code.</summary>
        Citation,
        /// <summary>Comments field code.</summary>
        Comments,
        /// <summary>Compare field code.</summary>
        Compare,
        /// <summary>CreateDate field code.</summary>
        CreateDate,
        /// <summary>Database field code.</summary>
        Database,
        /// <summary>Date field code.</summary>
        Date,
        /// <summary>DocProperty field code.</summary>
        DocProperty,
        /// <summary>DocVariable field code.</summary>
        DocVariable,
        /// <summary>Embed field code.</summary>
        Embed,
        /// <summary>FileName field code.</summary>
        FileName,
        /// <summary>FileSize field code.</summary>
        FileSize,
        /// <summary>GoToButton field code.</summary>
        GoToButton,
        /// <summary>GreetingLine field code.</summary>
        GreetingLine,
        /// <summary>HyperlinkIf field code.</summary>
        HyperlinkIf,
        /// <summary>IncludePicture field code.</summary>
        IncludePicture,
        /// <summary>IncludeText field code.</summary>
        IncludeText,
        /// <summary>Index field code.</summary>
        Index,
        /// <summary>Info field code.</summary>
        Info,
        /// <summary>Keywords field code.</summary>
        Keywords,
        /// <summary>LastSavedBy field code.</summary>
        LastSavedBy,
        /// <summary>Link field code.</summary>
        Link,
        /// <summary>ListNum field code.</summary>
        ListNum,
        /// <summary>MacroButton field code.</summary>
        MacroButton,
        /// <summary>MergeField field code.</summary>
        MergeField,
        /// <summary>MergeRec field code.</summary>
        MergeRec,
        /// <summary>MergeSeq field code.</summary>
        MergeSeq,
        /// <summary>Next field code.</summary>
        Next,
        /// <summary>NextIf field code.</summary>
        NextIf,
        /// <summary>NoteRef field code.</summary>
        NoteRef,
        /// <summary>NumChars field code.</summary>
        NumChars,
        /// <summary>NumPages field code.</summary>
        NumPages,
        /// <summary>NumWords field code.</summary>
        NumWords,
        /// <summary>Page field code.</summary>
        Page,
        /// <summary>PageRef field code.</summary>
        PageRef,
        /// <summary>Print field code.</summary>
        Print,
        /// <summary>PrintDate field code.</summary>
        PrintDate,
        /// <summary>Private field code.</summary>
        Private,
        /// <summary>Quote field code.</summary>
        Quote,
        /// <summary>RD field code.</summary>
        RD,
        /// <summary>Ref field code.</summary>
        Ref,
        /// <summary>RevNum field code.</summary>
        RevNum,
        /// <summary>SaveDate field code.</summary>
        SaveDate,
        /// <summary>Section field code.</summary>
        Section,
        /// <summary>SectionPages field code.</summary>
        SectionPages,
        /// <summary>Seq field code.</summary>
        Seq,
        /// <summary>Set field code.</summary>
        Set,
        /// <summary>SkipIf field code.</summary>
        SkipIf,
        /// <summary>StyleRef field code.</summary>
        StyleRef,
        /// <summary>Subject field code.</summary>
        Subject,
        /// <summary>Symbol field code.</summary>
        Symbol,
        /// <summary>TA field code.</summary>
        TA,
        /// <summary>TC field code.</summary>
        TC,
        /// <summary>Template field code.</summary>
        Template,
        /// <summary>Time field code.</summary>
        Time,
        /// <summary>Title field code.</summary>
        Title,
        /// <summary>TOA field code.</summary>
        TOA,
        /// <summary>TOC field code.</summary>
        TOC,
        /// <summary>UserAddress field code.</summary>
        UserAddress,
        /// <summary>UserInitials field code.</summary>
        UserInitials,
        /// <summary>UserName field code.</summary>
        UserName,
        XE
    }

    /// <summary>
    /// Specifies format switches for Word field codes.
    /// </summary>
    public enum WordFieldFormat {
        /// <summary>Lower format switch.</summary>
        Lower,
        /// <summary>Upper format switch.</summary>
        Upper,
        /// <summary>FirstCap format switch.</summary>
        FirstCap,
        /// <summary>Caps format switch.</summary>
        Caps,
        /// <summary>Mergeformat format switch.</summary>
        Mergeformat,
        /// <summary>Roman format switch.</summary>
        Roman,
        /// <summary>roman format switch.</summary>
        roman,
        /// <summary>Arabic format switch.</summary>
        Arabic,
        /// <summary>Alphabetical format switch.</summary>
        Alphabetical,
        /// <summary>ALPHABETICAL format switch.</summary>
        ALPHABETICAL,
        /// <summary>Ordinal format switch.</summary>
        Ordinal,
        /// <summary>OrdText format switch.</summary>
        OrdText,
        /// <summary>CardText format switch.</summary>
        CardText,
        /// <summary>DollarText format switch.</summary>
        DollarText,
        /// <summary>Hex format switch.</summary>
        Hex,
        /// <summary>CharFormat format switch.</summary>
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
