namespace OfficeIMO.Core.Internal.OpenXml;

internal sealed partial class OfficeXmlPullReader {
    private enum FrameKind : byte {
        Element,
        Transparent,
        AlternateContent
    }

    private struct Frame {
        internal int QNameStart;
        internal int QNameLength;
        internal int NamespaceId;
        internal int LocalStart;
        internal int LocalLength;
        internal int BindingMark;
        internal int IgnorableMark;
        internal int ProcessContentMark;
        internal FrameKind Kind;
        internal bool PreserveSpace;
        internal bool ChoiceSelected;
        internal bool SawCommentOrInstruction;
    }

    private struct Binding {
        internal int PrefixStart;
        internal int PrefixLength;
        internal int NamespaceId;
    }

    private struct ProcessContentEntry {
        internal int NamespaceId;
        // UTF-8 local name; null stands for every element of the namespace (prefix:*).
        internal byte[]? LocalName;
    }

    private struct Attribute {
        internal int PrefixStart;
        internal int PrefixLength;
        internal int LocalStart;
        internal int LocalLength;
        internal int ValueStart;
        internal int ValueLength;
        internal bool NeedsDecoding;
    }

    // Parses a start tag, pushes its frame and decides whether the caller sees it.
    private bool ReadStartTag() {
        byte[] data = _data;
        int qnameStart = _position + 1;
        int position = qnameStart;
        int colon = -1;
        while (position < _end && !IsNameTerminator(data[position])) {
            if (data[position] == (byte)':' && colon < 0) colon = position;
            position++;
        }
        int qnameLength = position - qnameStart;
        if (qnameLength == 0) throw Malformed("A start tag has no name.");

        _attributeCount = 0;
        bool empty = false;
        while (true) {
            while (position < _end && IsXmlSpace(data[position])) position++;
            if (position >= _end) throw Malformed("A start tag is not closed.");
            byte value = data[position];
            if (value == (byte)'>') {
                position++;
                break;
            }
            if (value == (byte)'/') {
                if (position + 1 >= _end || data[position + 1] != (byte)'>') throw Malformed("A start tag is not closed.");
                position += 2;
                empty = true;
                break;
            }
            position = ReadAttribute(position);
        }
        _position = position;

        if (_frameCount == 0) {
            // Counted before markup compatibility, so a root that compatibility processing hides still counts.
            if (_rootSeen) throw Malformed("The part has more than one root element.");
            _rootSeen = true;
        }
        if (_frameCount == MaximumDepth) throw Malformed("Elements are nested too deeply.");
        if (_frameCount == _frames.Length) Array.Resize(ref _frames, _frames.Length * 2);
        bool parentPreserves = _frameCount > 0 && _frames[_frameCount - 1].PreserveSpace;
        ref Frame frame = ref _frames[_frameCount++];
        frame = new Frame {
            QNameStart = qnameStart,
            QNameLength = qnameLength,
            LocalStart = colon < 0 ? qnameStart : colon + 1,
            LocalLength = colon < 0 ? qnameLength : qnameStart + qnameLength - colon - 1,
            BindingMark = _bindingCount,
            IgnorableMark = _ignorableCount,
            ProcessContentMark = _processContentCount,
            PreserveSpace = parentPreserves
        };

        DeclareNamespaces();
        frame.NamespaceId = colon < 0 ? ResolvePrefix(0, 0) : ResolvePrefix(qnameStart, colon - qnameStart);
        if (_attributeCount > 1) RejectDuplicateAttributes();
        if (_attributeCount > 0) ApplySpecialAttributes(ref frame);

        bool skip = false;
        frame.Kind = _processMarkupCompatibility ? Classify(ref frame, out skip) : FrameKind.Element;
        if (skip) {
            if (!empty) SkipContent();
            _frameCount--;
            _bindingCount = frame.BindingMark;
            _ignorableCount = frame.IgnorableMark;
            _processContentCount = frame.ProcessContentMark;
            return false;
        }
        if (frame.Kind != FrameKind.Element) {
            if (empty) PopFrame();
            return false;
        }

        _depth++;
        NodeKind = OfficeXmlNodeKind.StartElement;
        NamespaceId = frame.NamespaceId;
        _nameStart = frame.LocalStart;
        _nameLength = frame.LocalLength;
        IsEmptyElement = empty;
        IsCData = false;
        IsWhitespace = false;
        PreserveSpace = frame.PreserveSpace;
        _pendingEnd = empty;
        return true;
    }

    private int ReadAttribute(int position) {
        byte[] data = _data;
        int nameStart = position;
        int colon = -1;
        while (position < _end && data[position] != (byte)'=' && !IsXmlSpace(data[position])
            && data[position] != (byte)'>' && data[position] != (byte)'/') {
            if (data[position] == (byte)':' && colon < 0) colon = position;
            position++;
        }
        int nameEnd = position;
        if (nameEnd == nameStart) throw Malformed("An attribute has no name.");
        while (position < _end && IsXmlSpace(data[position])) position++;
        if (position >= _end || data[position] != (byte)'=') throw Malformed("An attribute has no value.");
        position++;
        while (position < _end && IsXmlSpace(data[position])) position++;
        if (position >= _end || data[position] != (byte)'"' && data[position] != (byte)'\'') {
            throw Malformed("An attribute value is not quoted.");
        }
        byte quote = data[position++];
        int valueStart = position;
        bool needsDecoding = false;
        while (position < _end && data[position] != quote) {
            byte value = data[position];
            if (value == (byte)'<') throw Malformed("An attribute value contains '<'.");
            if (value == (byte)'&' || value == (byte)'\r' || value == (byte)'\n' || value == (byte)'\t') needsDecoding = true;
            position++;
        }
        if (position >= _end) throw Malformed("An attribute value is not closed.");

        if (_attributeCount == _attributes.Length) Array.Resize(ref _attributes, _attributes.Length * 2);
        _attributes[_attributeCount++] = new Attribute {
            PrefixStart = nameStart,
            PrefixLength = colon < 0 ? 0 : colon - nameStart,
            LocalStart = colon < 0 ? nameStart : colon + 1,
            LocalLength = colon < 0 ? nameEnd - nameStart : nameEnd - colon - 1,
            ValueStart = valueStart,
            ValueLength = position - valueStart,
            NeedsDecoding = needsDecoding
        };
        return position + 1;
    }

    // Two attributes may not share an expanded name, even when different prefixes are bound to the same namespace.
    private void RejectDuplicateAttributes() {
        for (int first = 1; first < _attributeCount; first++) {
            ref Attribute later = ref _attributes[first];
            for (int second = 0; second < first; second++) {
                ref Attribute earlier = ref _attributes[second];
                if (earlier.LocalLength != later.LocalLength || !BytesEqual(earlier.LocalStart, later.LocalStart, later.LocalLength)) continue;
                if (AttributeNamespace(earlier) == AttributeNamespace(later)) throw Malformed("An element repeats an attribute.");
            }
        }
    }

    // Namespace declarations sit in the xmlns namespace; other unprefixed attributes have no namespace.
    private int AttributeNamespace(in Attribute attribute) {
        if (attribute.PrefixLength == 0) {
            return AsciiEquals(attribute.LocalStart, attribute.LocalLength, "xmlns") ? OfficeXmlNamespaces.Xmlns : OfficeXmlNamespaces.None;
        }
        if (attribute.PrefixLength == 5 && AsciiEquals(attribute.PrefixStart, 5, "xmlns")) return OfficeXmlNamespaces.Xmlns;
        return ResolvePrefix(attribute.PrefixStart, attribute.PrefixLength);
    }

    private void DeclareNamespaces() {
        for (int index = 0; index < _attributeCount; index++) {
            ref Attribute attribute = ref _attributes[index];
            int prefixStart;
            int prefixLength;
            if (attribute.PrefixLength == 0 && AsciiEquals(attribute.LocalStart, attribute.LocalLength, "xmlns")) {
                prefixStart = 0;
                prefixLength = 0;
            } else if (attribute.PrefixLength == 5 && AsciiEquals(attribute.PrefixStart, 5, "xmlns")) {
                prefixStart = attribute.LocalStart;
                prefixLength = attribute.LocalLength;
            } else {
                continue;
            }
            int namespaceId;
            if (attribute.ValueLength == 0) {
                if (prefixLength != 0) throw Malformed("A namespace prefix is bound to an empty URI.");
                namespaceId = OfficeXmlNamespaces.None;
            } else if (attribute.NeedsDecoding) {
                namespaceId = _namespaces.GetId(DecodeAttribute(attribute));
            } else {
                namespaceId = _namespaces.GetId(_data, attribute.ValueStart, attribute.ValueLength);
            }
            if (_bindingCount == _bindings.Length) Array.Resize(ref _bindings, _bindings.Length * 2);
            _bindings[_bindingCount++] = new Binding { PrefixStart = prefixStart, PrefixLength = prefixLength, NamespaceId = namespaceId };
        }
    }

    // Resolves a prefix (empty for the default namespace) against the bindings in scope.
    private int ResolvePrefix(int start, int length) {
        for (int index = _bindingCount - 1; index >= 0; index--) {
            ref Binding binding = ref _bindings[index];
            if (binding.PrefixLength == length && BytesEqual(binding.PrefixStart, start, length)) return binding.NamespaceId;
        }
        if (length == 0) return OfficeXmlNamespaces.None;
        if (length == 3 && AsciiEquals(start, 3, "xml")) return OfficeXmlNamespaces.Xml;
        throw Malformed("The prefix '" + System.Text.Encoding.UTF8.GetString(_data, start, length) + "' is not declared.");
    }

    private void ApplySpecialAttributes(ref Frame frame) {
        for (int index = 0; index < _attributeCount; index++) {
            ref Attribute attribute = ref _attributes[index];
            if (attribute.PrefixLength == 0 || attribute.PrefixLength == 5 && AsciiEquals(attribute.PrefixStart, 5, "xmlns")) continue;
            int namespaceId = ResolvePrefix(attribute.PrefixStart, attribute.PrefixLength);
            if (namespaceId == OfficeXmlNamespaces.Xml) {
                if (AsciiEquals(attribute.LocalStart, attribute.LocalLength, "space")) {
                    string value = GetAttributeValue(attribute);
                    if (value == "preserve") frame.PreserveSpace = true;
                    else if (value == "default") frame.PreserveSpace = false;
                }
            } else if (namespaceId == OfficeXmlNamespaces.MarkupCompatibility && _processMarkupCompatibility) {
                if (AsciiEquals(attribute.LocalStart, attribute.LocalLength, "Ignorable")) {
                    foreach (string prefix in SplitList(GetAttributeValue(attribute))) {
                        // A prefix with no binding cannot name any element or attribute, so it has nothing to ignore.
                        if (!TryResolvePrefix(prefix, out int ignorable)) continue;
                        if (_ignorableCount == _ignorable.Length) Array.Resize(ref _ignorable, Math.Max(4, _ignorable.Length * 2));
                        _ignorable[_ignorableCount++] = ignorable;
                    }
                } else if (AsciiEquals(attribute.LocalStart, attribute.LocalLength, "MustUnderstand")) {
                    foreach (string prefix in SplitList(GetAttributeValue(attribute))) {
                        if (!TryResolvePrefix(prefix, out int required) || !IsUnderstood(required)) {
                            throw new InvalidDataException("The part requires the namespace of prefix '" + prefix + "', which this reader does not understand.");
                        }
                    }
                } else if (AsciiEquals(attribute.LocalStart, attribute.LocalLength, "ProcessContent")) {
                    foreach (string name in SplitList(GetAttributeValue(attribute))) {
                        int colon = name.IndexOf(':');
                        if (colon <= 0 || !TryResolvePrefix(name.Substring(0, colon), out int processed)) continue;
                        string local = name.Substring(colon + 1);
                        if (_processContentCount == _processContent.Length) Array.Resize(ref _processContent, Math.Max(4, _processContent.Length * 2));
                        _processContent[_processContentCount++] = new ProcessContentEntry {
                            NamespaceId = processed,
                            LocalName = local == "*" ? null : System.Text.Encoding.UTF8.GetBytes(local)
                        };
                    }
                }
            }
        }
    }

    private FrameKind Classify(ref Frame frame, out bool skip) {
        skip = false;
        if (_frameCount > 1) {
            ref Frame parent = ref _frames[_frameCount - 2];
            if (parent.Kind == FrameKind.AlternateContent) {
                bool selected = false;
                if (frame.NamespaceId == OfficeXmlNamespaces.MarkupCompatibility) {
                    if (AsciiEquals(frame.LocalStart, frame.LocalLength, "Choice")) {
                        selected = !parent.ChoiceSelected && RequirementsUnderstood();
                    } else if (AsciiEquals(frame.LocalStart, frame.LocalLength, "Fallback")) {
                        selected = !parent.ChoiceSelected;
                    }
                }
                if (selected) {
                    parent.ChoiceSelected = true;
                    return FrameKind.Transparent;
                }
                skip = true;
                return FrameKind.Transparent;
            }
        }
        if (frame.NamespaceId == OfficeXmlNamespaces.MarkupCompatibility
            && AsciiEquals(frame.LocalStart, frame.LocalLength, "AlternateContent")) {
            return FrameKind.AlternateContent;
        }
        if (!IsUnderstood(frame.NamespaceId) && IsIgnorable(frame.NamespaceId)) {
            if (ProcessesContent(frame.NamespaceId, frame.LocalStart, frame.LocalLength)) return FrameKind.Transparent;
            skip = true;
            return FrameKind.Transparent;
        }
        return FrameKind.Element;
    }

    private bool RequirementsUnderstood() {
        for (int index = 0; index < _attributeCount; index++) {
            ref Attribute attribute = ref _attributes[index];
            if (attribute.PrefixLength != 0 || !AsciiEquals(attribute.LocalStart, attribute.LocalLength, "Requires")) continue;
            foreach (string prefix in SplitList(GetAttributeValue(attribute))) {
                if (!TryResolvePrefix(prefix, out int required) || !IsUnderstood(required)) return false;
            }
            return true;
        }
        throw Malformed("mc:Choice has no Requires attribute.");
    }

    // The reader understands the namespaces of ECMA-376 first edition, as an Office 2007 consumer does.
    private static bool IsUnderstood(int namespaceId) => OfficeXmlNamespaces.IsEcma376(namespaceId);

    private bool IsIgnorable(int namespaceId) {
        for (int index = 0; index < _ignorableCount; index++) {
            if (_ignorable[index] == namespaceId) return true;
        }
        return false;
    }

    private bool ProcessesContent(int namespaceId, int localStart, int localLength) {
        for (int index = 0; index < _processContentCount; index++) {
            ref ProcessContentEntry entry = ref _processContent[index];
            if (entry.NamespaceId == namespaceId && (entry.LocalName == null || Utf8Equals(localStart, localLength, entry.LocalName))) return true;
        }
        return false;
    }

    private bool TryResolvePrefix(string prefix, out int namespaceId) {
        byte[] bytes = System.Text.Encoding.UTF8.GetBytes(prefix);
        for (int index = _bindingCount - 1; index >= 0; index--) {
            ref Binding binding = ref _bindings[index];
            if (binding.PrefixLength != 0 && Utf8Equals(binding.PrefixStart, binding.PrefixLength, bytes)) {
                namespaceId = binding.NamespaceId;
                return true;
            }
        }
        namespaceId = prefix == "xml" ? OfficeXmlNamespaces.Xml : OfficeXmlNamespaces.None;
        return prefix == "xml";
    }

    private static string[] SplitList(string value) =>
        value.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

    // Skips everything up to and including the end tag of the element whose start tag was just read.
    private void SkipContent() {
        int depth = 1;
        byte[] data = _data;
        while (true) {
            int position = _position;
            while (position < _end && data[position] != (byte)'<') position++;
            if (position + 1 >= _end) throw Malformed("The part ends inside a skipped element.");
            _position = position;
            byte next = data[position + 1];
            if (next == (byte)'/') {
                SkipPast(">", 2);
                if (--depth == 0) return;
            } else if (next == (byte)'?') {
                SkipPast("?>", 2);
            } else if (next == (byte)'!') {
                if (StartsWith(position, "<!--")) SkipPast("-->", 4);
                else if (StartsWith(position, "<![CDATA[")) SkipPast("]]>", 9);
                else throw Malformed("The part has a document type declaration.");
            } else {
                // Attribute values may contain '>', so the start tag is scanned quote-aware.
                position += 1;
                byte quote = 0;
                while (position < _end) {
                    byte value = data[position];
                    if (quote != 0) {
                        if (value == quote) quote = 0;
                    } else if (value == (byte)'"' || value == (byte)'\'') {
                        quote = value;
                    } else if (value == (byte)'>') {
                        break;
                    }
                    position++;
                }
                if (position >= _end) throw Malformed("A start tag is not closed.");
                if (data[position - 1] != (byte)'/') depth++;
                _position = position + 1;
            }
        }
    }
}
