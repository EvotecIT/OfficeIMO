using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word.OpenXml;
using System.Text;

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
        /// <summary>Index entry field code.</summary>
        XE,
        /// <summary>Formula field code, represented in Word field instructions as an equals expression.</summary>
        Formula,
        /// <summary>EQ mathematical equation field code used by legacy Word and RTF.</summary>
        EQ
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
        MergeFormat,
        /// <summary>Roman format switch.</summary>
        UpperRoman,
        /// <summary>roman format switch.</summary>
        LowerRoman,
        /// <summary>Arabic format switch.</summary>
        Arabic,
        /// <summary>Alphabetical format switch.</summary>
        LowerLetter,
        /// <summary>ALPHABETICAL format switch.</summary>
        UpperLetter,
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
        private const int W = OfficeXmlNamespaces.WordprocessingML;
        private WordParagraph? _authoringParagraph;

        /// <summary>Gets the containing paragraph handle for continuing paragraph authoring.</summary>
        public WordParagraph Paragraph {
            get {
                if (_authoringParagraph != null) return _authoringParagraph;
                ResolveNodes();
                return WordParagraph.WrapSdkField(_document, _paragraph, _simpleField, _runs);
            }
        }

        private readonly WordDocument _document;
        private OfficeXmlNode _paragraph;
        private OfficeXmlNode _simpleField;
        private readonly List<OfficeXmlNode> _runs;
        private readonly Action<OfficeXmlNode>? _replaceSdkRuns;

        /// <summary>
        /// Gets the type of the current field.
        /// </summary>
        public WordFieldType? FieldType => GetParsedField().WordFieldType;

        /// <summary>
        /// Gets the format switches applied to the field.
        /// </summary>
        public IReadOnlyList<WordFieldFormat> FieldFormat => new List<WordFieldFormat>(GetParsedField().FormatSwitches);

        /// <summary>
        /// Gets the raw switch parameters from the field code.
        /// </summary>
        public List<String> FieldSwitches => new List<String>(GetParsedField().Switches);

        /// <summary>
        /// Gets the instructions portion of the field code.
        /// </summary>
        public List<String> FieldInstructions => new List<String>(GetParsedField().Instructions);

        private string? _parsedFieldCode;
        private WordFieldParser? _parsedField;

        // Reading several properties of one field would otherwise parse the same code each time.
        // The cache is keyed by the current field code, so edits to the instruction are picked up.
        private WordFieldParser GetParsedField() {
            string fieldCode = Field;
            WordFieldParser? parsed = _parsedField;
            if (parsed == null || !string.Equals(_parsedFieldCode, fieldCode, StringComparison.Ordinal)) {
                parsed = new WordFieldParser(fieldCode);
                _parsedField = parsed;
                _parsedFieldCode = fieldCode;
            }
            return parsed;
        }

        /// <summary>
        /// Gets the raw field code.
        /// </summary>
        public string Field {
            get {
                ResolveNodes();
                if (_simpleField.Exists) return _simpleField.Attribute(W, "instr") ?? string.Empty;
                var instruction = new StringBuilder();
                foreach (OfficeXmlNode run in _runs) instruction.Append(run.Child(W, "instrText").InnerText);
                return instruction.ToString();
            }
        }

        /// <summary>Gets or sets a value indicating whether the field is marked dirty.</summary>
        public bool UpdateField {
            get => GetFlag("dirty");
            set => SetFlag("dirty", value);
        }

        /// <summary>Gets or sets a value indicating whether the field is locked.</summary>
        public bool LockField {
            get => GetFlag("fldLock");
            set => SetFlag("fldLock", value);
        }

        private bool GetFlag(string attribute) {
            ResolveNodes();
            if (_simpleField.Exists) return WordFieldInventory.ReadFlag(_simpleField, attribute);
            foreach (OfficeXmlNode run in _runs) {
                OfficeXmlNode marker = run.Child(W, "fldChar");
                if (marker.Exists) return WordFieldInventory.ReadFlag(marker, attribute);
            }
            return false;
        }

        private void SetFlag(string attribute, bool value) {
            ResolveNodes()?.MarkChanged();
            string token = value ? "true" : "false";
            if (_simpleField.Exists) {
                _simpleField.SetAttribute(W, attribute, token);
                return;
            }
            foreach (OfficeXmlNode run in _runs) {
                OfficeXmlNode marker = run.Child(W, "fldChar");
                if (marker.Exists) marker.SetAttribute(W, attribute, token);
            }
        }

        /// <summary>Gets or sets the display text of the field.</summary>
        public string Text {
            get {
                ResolveNodes();
                var text = new StringBuilder();
                foreach (OfficeXmlNode run in _runs) {
                    foreach (OfficeXmlNode item in run.Children(W, "t")) text.Append(item.InnerText);
                }
                return text.ToString();
            }
            set {
                WordPackageModel? model = ResolveNodes();
                bool first = true;
                foreach (OfficeXmlNode run in _runs) {
                    foreach (OfficeXmlNode text in run.Children(W, "t")) {
                        if (first) {
                            text.SetInnerText(value ?? string.Empty);
                            text.SetAttribute(OfficeXmlNamespaces.Xml, "space", "preserve");
                            model?.MarkChanged();
                            first = false;
                        } else {
                            text.SetInnerText(string.Empty);
                        }
                    }
                }
            }
        }

        internal WordField(WordDocument document, OfficeXmlNode paragraph, OfficeXmlNode simpleField,
            List<OfficeXmlNode>? runs = null, Action<OfficeXmlNode>? replaceSdkRuns = null) {
            _document = document;
            _paragraph = paragraph;
            _simpleField = simpleField;
            _runs = runs ?? simpleField.Children(W, "r").ToList();
            _replaceSdkRuns = replaceSdkRuns;
        }

        // Field wrappers may outlive the owned package, just like their containing paragraphs.
        private WordPackageModel? ResolveNodes() {
            if (_paragraph.Tree is not { } tree) return null;
            if (_document._ownedModel is { } model && _document.OwnsWordTree(model, tree)) return model;

            WordPackageModel? copy = _document.DisposedReadModel;
            OfficeXmlTree? copyTree = copy == null || _document._ownedNodeElements != null
                ? null : _document.GetDisposedWordTree(copy, tree);
            if (copyTree != null) {
                if (!ReferenceEquals(copyTree, tree)) {
                    _paragraph = CopyOf(_paragraph);
                    _simpleField = CopyOf(_simpleField);
                    for (int index = 0; index < _runs.Count; index++) _runs[index] = CopyOf(_runs[index]);
                }
                return copy;
            }

            if (_document.FindSdkWordNodeMap(tree) is not { } elements) {
                throw new InvalidOperationException("The field is no longer part of the document.");
            }
            _paragraph = SdkNode(_paragraph);
            _simpleField = SdkNode(_simpleField);
            for (int index = 0; index < _runs.Count; index++) _runs[index] = SdkNode(_runs[index]);
            return null;

            OfficeXmlNode CopyOf(OfficeXmlNode node) => node.IsNone
                ? default : new OfficeXmlNode(copyTree!, _document.DisposedCopyNode(tree, node.Node));

            OfficeXmlNode SdkNode(OfficeXmlNode node) => node.Node >= 0 && node.Node < elements.Length
                ? OfficeXmlNode.FromSdk(elements[node.Node]) : default;
        }

        internal void SetAuthoringParagraph(WordParagraph paragraph) => _authoringParagraph = paragraph;

        internal (OfficeXmlNode SimpleField, List<OfficeXmlNode> Runs) CurrentNodes() {
            ResolveNodes();
            return (_simpleField, _runs);
        }
    }
}
