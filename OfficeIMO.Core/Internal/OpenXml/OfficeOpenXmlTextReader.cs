namespace OfficeIMO.Core.Internal.OpenXml;

/// <summary>
/// Plain-text reads of Word and PowerPoint packages on <see cref="OfficeOpcPackage"/> and
/// <see cref="OfficeXmlPullReader"/>, without the Open XML SDK.
/// </summary>
/// <remarks>
/// The text of an element is the text of its leaf descendants in document order, as the Open XML SDK reports
/// <c>InnerText</c>: character data directly inside an element that has child elements is ignored, a leaf's text ends
/// at its first comment or processing instruction, and whitespace-only character data counts in the text elements
/// (<c>w:t</c>, <c>w:delText</c>, <c>w:instrText</c>, <c>w:delInstrText</c>, <c>m:t</c>, <c>a:t</c>) but elsewhere
/// only inside an <c>xml:space="preserve"</c> scope or a CDATA section. Markup compatibility is applied as an ECMA-376
/// first-edition consumer, so <c>mc:AlternateContent</c> contributes the text of one branch.
/// </remarks>
internal static class OfficeOpenXmlTextReader {
    /// <summary>Returns the text of each paragraph that is a direct child of the main document body.</summary>
    internal static IEnumerable<string> ReadWordBodyParagraphs(Stream package) {
        using OfficeOpcPackage opc = OfficeOpcPackage.Open(package, leaveOpen: true);
        string mainPart = opc.GetRelatedPartName(null, OfficeOpcPackage.OfficeDocumentRelationshipType)
            ?? throw new InvalidDataException("The package has no main document part.");
        OfficeXmlPullReader reader = opc.OpenXml(mainPart, out OfficePooledBytes content);
        using (content) {
            var text = new OfficeXmlInnerText();
            bool inBody = false;
            while (reader.Read()) {
                if (reader.NodeKind == OfficeXmlNodeKind.StartElement) {
                    if (reader.Depth == 1) {
                        if (!reader.IsElement(OfficeXmlNamespaces.WordprocessingML, "document")) {
                            throw new InvalidDataException("The main document part does not start with w:document.");
                        }
                    } else if (reader.Depth == 2) {
                        inBody = reader.IsElement(OfficeXmlNamespaces.WordprocessingML, "body");
                        if (!inBody) reader.SkipElement();
                    } else if (reader.Depth == 3 && inBody) {
                        if (reader.IsElement(OfficeXmlNamespaces.WordprocessingML, "p")) {
                            yield return text.Read(reader);
                        } else {
                            reader.SkipElement();
                        }
                    }
                }
            }
        }
    }

    /// <summary>Returns the text of every DrawingML paragraph on each slide, slides in presentation order.</summary>
    internal static IEnumerable<OfficeSlideParagraph> ReadSlideParagraphs(Stream package) {
        using OfficeOpcPackage opc = OfficeOpcPackage.Open(package, leaveOpen: true);
        string presentationPart = opc.GetRelatedPartName(null, OfficeOpcPackage.OfficeDocumentRelationshipType)
            ?? throw new InvalidDataException("The package has no presentation part.");
        List<string> slideParts = ReadSlideOrder(opc, presentationPart);
        var text = new OfficeXmlInnerText();
        for (int slide = 0; slide < slideParts.Count; slide++) {
            OfficeXmlPullReader reader = opc.OpenXml(slideParts[slide], out OfficePooledBytes content);
            using (content) {
                while (reader.Read()) {
                    if (reader.NodeKind == OfficeXmlNodeKind.StartElement && reader.IsElement(OfficeXmlNamespaces.DrawingML, "p")) {
                        yield return new OfficeSlideParagraph(slide, text.Read(reader));
                    }
                }
            }
        }
    }

    private static List<string> ReadSlideOrder(OfficeOpcPackage opc, string presentationPart) {
        var relationshipIds = new List<string>();
        OfficeXmlPullReader reader = opc.OpenXml(presentationPart, out OfficePooledBytes content);
        using (content) {
            while (reader.Read()) {
                if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
                if (reader.Depth == 2 && !reader.IsElement(OfficeXmlNamespaces.PresentationML, "sldIdLst")) {
                    reader.SkipElement();
                } else if (reader.Depth == 3 && reader.IsElement(OfficeXmlNamespaces.PresentationML, "sldId")
                    && reader.TryGetAttribute(OfficeXmlNamespaces.Relationships, "id", out string id)) {
                    relationshipIds.Add(id);
                }
            }
        }
        IReadOnlyList<OfficeOpcRelationship> relationships = opc.GetRelationships(presentationPart);
        var byId = new Dictionary<string, OfficeOpcRelationship>(relationships.Count, StringComparer.Ordinal);
        foreach (OfficeOpcRelationship relationship in relationships) byId[relationship.Id] = relationship;
        var slideParts = new List<string>(relationshipIds.Count);
        foreach (string id in relationshipIds) {
            if (!byId.TryGetValue(id, out OfficeOpcRelationship? relationship) || relationship.TargetPartName == null
                || !string.Equals(relationship.Type, OfficeOpcPackage.SlideRelationshipType, StringComparison.Ordinal)) {
                throw new InvalidDataException("The slide list refers to '" + id + "', which is not a slide relationship.");
            }
            slideParts.Add(relationship.TargetPartName);
        }
        return slideParts;
    }
}

/// <summary>The text of one paragraph and the zero-based position of its slide in the presentation.</summary>
internal readonly struct OfficeSlideParagraph {
    internal OfficeSlideParagraph(int slideIndex, string text) {
        SlideIndex = slideIndex;
        Text = text;
    }

    internal int SlideIndex { get; }
    internal string Text { get; }
}

/// <summary>Collects the leaf text of the element the reader is positioned on.</summary>
internal sealed class OfficeXmlInnerText {
    private readonly OfficeCharBuffer _buffer = new();
    private Level[] _levels = new Level[16];

    private struct Level {
        internal int Mark;
        internal bool HasChild;
        internal bool IsTextElement;
    }

    /// <summary>Reads to the end of the current element and returns its text.</summary>
    internal string Read(OfficeXmlPullReader reader) {
        int rootDepth = reader.Depth;
        _buffer.Clear();
        _levels[0] = new Level { IsTextElement = IsTextElement(reader) };
        int top = 0;
        if (reader.IsEmptyElement) {
            reader.Read();
            return string.Empty;
        }
        while (reader.Read()) {
            switch (reader.NodeKind) {
                case OfficeXmlNodeKind.StartElement:
                    ref Level parent = ref _levels[top];
                    if (!parent.HasChild) {
                        // An element with children is not a text leaf, so its own character data does not count.
                        _buffer.Truncate(parent.Mark);
                        parent.HasChild = true;
                    }
                    if (++top == _levels.Length) Array.Resize(ref _levels, _levels.Length * 2);
                    _levels[top] = new Level { Mark = _buffer.Length, IsTextElement = IsTextElement(reader) };
                    break;
                case OfficeXmlNodeKind.EndElement:
                    if (reader.Depth == rootDepth) return _buffer.ToString();
                    top--;
                    break;
                case OfficeXmlNodeKind.Text:
                    ref Level level = ref _levels[top];
                    if (!level.HasChild && !reader.FollowsCommentOrInstruction
                        && (!reader.IsWhitespace || level.IsTextElement || reader.PreserveSpace || reader.IsCData)) {
                        reader.AppendText(_buffer);
                    }
                    break;
            }
        }
        throw new InvalidDataException("The part ends inside an element.");
    }

    private static bool IsTextElement(OfficeXmlPullReader reader) => reader.NamespaceId switch {
        OfficeXmlNamespaces.WordprocessingML => reader.LocalNameEquals("t") || reader.LocalNameEquals("delText")
            || reader.LocalNameEquals("instrText") || reader.LocalNameEquals("delInstrText"),
        OfficeXmlNamespaces.Math or OfficeXmlNamespaces.DrawingML => reader.LocalNameEquals("t"),
        _ => false
    };
}
