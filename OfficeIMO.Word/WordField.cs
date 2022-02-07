using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public class WordField {
        private readonly WordDocument _document;
        private readonly Paragraph _paragraph;
        private readonly List<Run> _runs = new List<Run>();
        private readonly SimpleField _simpleField;

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
    }
}
