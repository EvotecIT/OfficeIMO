using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>
    /// Renders consecutive paragraphs that share one <c>w:framePr</c> as a positioned frame.
    /// Word lays a text frame out like a borderless one-cell floating table: the same anchors,
    /// alignment, offsets, text clearance and wrapping. The frame takes no vertical flow
    /// unless it wraps as <c>notBeside</c>.
    /// </summary>
    private static bool TryRenderNativeParagraphFrame(INativePdfFlow pdf, IReadOnlyList<WordElement> elements, ref int index,
        Func<WordParagraph, (int Level, string Marker)?> getMarker, NativeNoteNumbering notes, WordPdfExportOptions? options,
        double? contentWidth, NativeDocumentDefaults defaults, NativeFontMap fontMap,
        IReadOnlyDictionary<OfficeXmlNode, string>? headingDestinations = null) {
        if (elements[index] is not WordParagraph first || !CanRenderNativeParagraphFrame(first) ||
            ResolveNativeParagraphFrame(first) is not { Exists: true } frame) return false;
        bool notBeside = frame.WAttr("wrap") == "notBeside";
        if (!notBeside && !pdf.SupportsPositionedTables) return false;

        var paragraphs = new List<WordParagraph> { first };
        string key = NativeFrameMarkup(frame);
        while (index + paragraphs.Count < elements.Count && elements[index + paragraphs.Count] is WordParagraph next &&
               CanRenderNativeParagraphFrame(next) && ResolveNativeParagraphFrame(next) is { Exists: true } nextFrame &&
               NativeFrameMarkup(nextFrame) == key) paragraphs.Add(next);

        WordTable table = CreateNativeFrameTable(first._document, frame, paragraphs, notBeside);
        foreach (WordParagraph paragraph in paragraphs) {
            OfficeXmlNode node = paragraph.Node();
            if (node.Exists) options?.DiagnosedFrameParagraphs.Add(node);
        }

        // The frame content is a copy: list markers, numbering and story relationships still
        // resolve through the source paragraphs.
        var sources = new Dictionary<OfficeXmlNode, OfficeXmlNode>();
        List<OfficeXmlNode> directCopies = table.Node().ModelDescendants("tc").FirstOrDefault().WElements("p").ToList();
        for (int i = 0; i < paragraphs.Count && i < directCopies.Count; i++) {
            OfficeXmlNode source = paragraphs[i].Node();
            sources[directCopies[i]] = source;
            List<OfficeXmlNode> nestedCopies = directCopies[i].ModelDescendants("p").ToList();
            List<OfficeXmlNode> nestedSources = source.ModelDescendants("p").ToList();
            for (int nested = 0; nested < nestedCopies.Count && nested < nestedSources.Count; nested++)
                sources[nestedCopies[nested]] = nestedSources[nested];
        }
        List<OfficeXmlNode> copies = sources.Keys.ToList();
        Func<WordParagraph, (int Level, string Marker)?> frameMarker = paragraph =>
            paragraph.Node() is { Exists: true } copy && sources.TryGetValue(copy, out OfficeXmlNode source)
                ? getMarker(WordConverterNodes.ParagraphView(first._document, source)) : getMarker(paragraph);

        if (defaults.RunningContentContext == null) {
            // Bookmarks in the frame stay link targets at the frame's position.
            foreach (WordParagraph paragraph in paragraphs) {
                RenderNativeParagraphBookmarks(pdf, paragraph);
                OfficeXmlNode node = paragraph.Node();
                if (node.Exists && !GetNativeParagraphBookmarkNames(paragraph).Any() &&
                    headingDestinations != null && headingDestinations.TryGetValue(node, out string? destination))
                    pdf.Bookmark(destination);
            }
        }

        var collapseFlow = pdf as NativeSpacingCollapseFlow;
        var collapseState = collapseFlow?.CaptureState();
        PdfCore.TableBlock? block;
        // The copies are diagnosed through their sources; their entries last only while the frame is built.
        foreach (OfficeXmlNode copy in copies) options?.DiagnosedFrameParagraphs.Add(copy);
        try {
            using (UseNativeFrameSources(sources))
            using (WordConverterNodes.UseListParagraphAliases(first._document, sources))
            using (WordConverterNodes.UseDetachedOwner(table.Node(), first._document, first.Node()))
                block = BuildNativeTableBlock(table, frameMarker, notes, options, contentWidth, defaults, fontMap, pdf.SupportsPositionedTables);
        } finally {
            foreach (OfficeXmlNode copy in copies) options?.DiagnosedFrameParagraphs.Remove(copy);
        }
        if (block != null) {
            if (block.Style is { Position: { } position } blockStyle) {
                blockStyle.Position = CreateNativeFramePosition(position, frame);
                // A drop cap moves to the next page with its paragraph.
                if (IsNativeDropCapFrame(frame)) blockStyle.KeepWithNext = true;
            }
            pdf.Table(block.Cells.Select(row => row.ToArray()), block.Align, block.Style);
        }
        if (collapseFlow != null && collapseState is { } state) {
            if (!notBeside) {
                // A floating frame leaves the flow untouched: the next paragraph collapses its
                // space before against the paragraph above the frame. At a section start Word
                // drops that paragraph's space before, as the frame starts the page.
                collapseFlow.RestorePendingSpacingAfter(state.Pending);
            } else if (state.Preserve) {
                // A notBeside frame at a section start passes the carried spacing on to the
                // paragraph after it, which does not collapse with the frame's space after.
                collapseFlow.RestoreState(state);
            } else {
                // Elsewhere the next paragraph collapses with the frame's last space after.
                collapseFlow.RestorePendingSpacingAfter(CreateNativeParagraphStyle(paragraphs[paragraphs.Count - 1], defaults, fontMap).SpacingAfter?.ToPoints());
            }
        }
        index += paragraphs.Count - 1;
        return true;
    }

    /// <summary>The frame's markup, which identifies paragraphs that share one frame.</summary>
    private static string NativeFrameMarkup(OfficeXmlNode frame) => Encoding.UTF8.GetString(frame.OuterXml());

    /// <summary>
    /// Frames are placed for paragraphs directly in the body or in a header or footer whose
    /// content the frame projection draws in full: text, fields, inline pictures and form
    /// controls. Frames in table cells, notes, text boxes and multi-column sections, frames that
    /// start a page or hold a page or column break, body headings, and frames with paragraph
    /// shading, borders, charts, SmartArt, shapes, text boxes, floating pictures, embedded
    /// objects, equations, note references or page-number fields in the body keep the
    /// ordinary-paragraph rendering, which already draws those. A drop cap is a frame at the
    /// start of the following paragraph: the text beside it wraps for the cap's lines, and a
    /// margin drop cap hangs in the margin beside the text.
    /// </summary>
    private static bool CanRenderNativeParagraphFrame(WordParagraph paragraph) {
        OfficeXmlNode source = paragraph.Node();
        OfficeXmlNode parent = source.IsNone ? default : source.Parent;
        bool inBody = parent.IsW("body");
        if (!(inBody || parent.IsW("hdr") || parent.IsW("ftr")) || ResolveNativeParagraphFrame(paragraph).IsNone) return false;
        if (source.DescendantElements(WordXml.W, "br").Any(IsNativeFlowBreak))
            return false;
        if (!IsNativeFrameContextSupported(paragraph)) return false;
        if (HasNativePageBreakBefore(paragraph) || GetNativeEffectiveParagraphShadingFill(paragraph) != null ||
            HasNativeFrameParagraphBorders(paragraph) || GetNativeRuns(paragraph).Any(run => run.Chart != null) ||
            HasNativeFrameUnsupportedContent(source, inBody)) return false;
        // A framed heading keeps the heading route, which also owns its outline entry.
        if (inBody && GetHeadingLevel(paragraph) > 0) return false;
        return !inBody || (paragraph.GetBodySection()?.ColumnCount ?? 1) <= 1;
    }

    private const string NativeFramePictureGraphicUri = "http://schemas.openxmlformats.org/drawingml/2006/picture";
    private static readonly string[] NativeFramePageFieldNames = { "PAGE", "NUMPAGES", "SECTIONPAGES", "SECTION", "PAGEREF" };

    private static bool HasNativeFrameUnsupportedContent(OfficeXmlNode paragraph, bool inBody) {
        if (inBody && HasNativeFramePageComplexField(paragraph)) return true;
        foreach (OfficeXmlNode element in paragraph.DescendantElements()) {
            if (element.Is(WordXml.Wp, "anchor") || element.Is(WordXml.Mc, "AlternateContent") ||
                element.Is(WordXml.M, "oMath") || element.Is(WordXml.M, "oMathPara"))
                return true;
            if (element.Is(WordXml.W, "pict") || element.Is(WordXml.W, "object") || element.Is(WordXml.W, "txbxContent") ||
                element.Is(WordXml.W, "footnoteReference") || element.Is(WordXml.W, "endnoteReference"))
                return true;
            if (element.Is(WordXml.A, "graphicData") && element.Attr("uri") != NativeFramePictureGraphicUri)
                return true;
            if (inBody && element.Is(WordXml.W, "fldSimple") && IsNativeFramePageField(element.WAttr("instr")))
                return true;
        }
        return false;
    }

    /// <summary>
    /// Reads each complex field's instruction as a whole, since Word often splits it across
    /// runs (" PA" + "GE "), including nested fields, and reports whether one paginates.
    /// </summary>
    private static bool HasNativeFramePageComplexField(OfficeXmlNode paragraph) {
        var instructions = new Stack<System.Text.StringBuilder?>();
        foreach (OfficeXmlNode element in paragraph.DescendantElements()) {
            if (element.IsW("fldChar")) {
                string? type = element.WAttr("fldCharType");
                if (type == "begin") {
                    instructions.Push(new System.Text.StringBuilder());
                } else if (instructions.Count > 0 && (type == "separate" || type == "end")) {
                    System.Text.StringBuilder? instruction = instructions.Pop();
                    if (instruction != null && IsNativeFramePageField(instruction.ToString())) return true;
                    // After the separator the field shows its result until its end mark.
                    if (type == "separate") instructions.Push(null);
                }
            } else if (element.IsW("instrText") && instructions.Count > 0 && instructions.Peek() is { } current) {
                current.Append(element.InnerText);
            }
        }
        return false;
    }

    private static bool IsNativeFramePageField(string? instruction) {
        string name = (instruction ?? string.Empty).Trim().Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? string.Empty;
        return NativeFramePageFieldNames.Any(field => string.Equals(field, name, StringComparison.OrdinalIgnoreCase));
    }

    private static readonly System.Threading.AsyncLocal<IReadOnlyDictionary<OfficeXmlNode, OfficeXmlNode>?> NativeFrameSources = new();

    private sealed class NativeFrameSourceScope : IDisposable {
        private readonly IReadOnlyDictionary<OfficeXmlNode, OfficeXmlNode>? _previous;
        internal NativeFrameSourceScope(IReadOnlyDictionary<OfficeXmlNode, OfficeXmlNode>? previous) => _previous = previous;
        public void Dispose() => NativeFrameSources.Value = _previous;
    }

    private static IDisposable UseNativeFrameSources(IReadOnlyDictionary<OfficeXmlNode, OfficeXmlNode> sources) {
        IReadOnlyDictionary<OfficeXmlNode, OfficeXmlNode>? previous = NativeFrameSources.Value;
        NativeFrameSources.Value = sources;
        return new NativeFrameSourceScope(previous);
    }

    /// <summary>The source paragraph of a projected frame copy, so section metrics such as the line grid still apply.</summary>
    private static WordParagraph ResolveNativeFrameSourceParagraph(WordParagraph paragraph) =>
        paragraph.Node() is { Exists: true } copy && NativeFrameSources.Value is { } sources && sources.TryGetValue(copy, out OfficeXmlNode source)
            ? WordConverterNodes.ParagraphView(paragraph._document, source) : paragraph;

    private static bool IsNativeFrameCopy(WordParagraph paragraph) =>
        paragraph.Node() is { Exists: true } copy && NativeFrameSources.Value?.ContainsKey(copy) == true;

    /// <summary>
    /// Context the frame copy cannot reproduce keeps the ordinary paragraph route: section
    /// marks and line numbering, right-to-left text, keep and outline settings, automatic spacing, tracked
    /// changes, hidden text, content controls and custom markup, fields that span paragraphs,
    /// and relationship-backed links in headers and footers.
    /// </summary>
    private static bool IsNativeFrameContextSupported(WordParagraph paragraph) {
        OfficeXmlNode source = paragraph.Node();
        OfficeXmlNode properties = source.WChild("pPr");
        if (IsSectionMarkOnly(source) || properties.WChild("sectPr").Exists) return false;
        NativeParagraphStyleDefaults styles = GetNativeParagraphStyleDefaults(paragraph);
        // Word writes keep-with-next on every drop cap; the cap is drawn beside the next paragraph.
        bool dropCapFrame = IsNativeDropCapFrame(ResolveNativeParagraphFrame(paragraph));
        if (properties.WChild("bidi").OnOff() == true || (!dropCapFrame && properties.WChild("keepNext").OnOff() == true) ||
            properties.WChild("keepLines").OnOff() == true ||
            properties.WChild("outlineLvl").Exists || styles.KeepWithNext == true || styles.KeepTogether == true) return false;
        if (GetNativeParagraphStyleChain(paragraph._document, paragraph.StyleId).Any(style =>
                style.WChild("pPr") is var styleParagraph &&
                (styleParagraph.WChild("bidi").OnOff() == true || styleParagraph.WChild("outlineLvl").Exists ||
                 style.WChild("rPr").WChild("rtl").OnOff() == true))) return false;
        // Automatic spacing follows table-cell rules inside the projection, which Word does not apply to frames.
        NativeParagraphSpacing spacing = styles.ParagraphSpacing.Merge(properties);
        OfficeXmlNode defaultSpacing = paragraph._document.DocumentDefaultsNode().WChild("pPrDefault").WChild("pPr").WChild("spacing");
        if (spacing.BeforeAuto == true || spacing.AfterAuto == true ||
            (spacing.BeforeAuto == null && WordXml.ParseOnOff(defaultSpacing.WAttr("beforeAutospacing")) == true) ||
            (spacing.AfterAuto == null && WordXml.ParseOnOff(defaultSpacing.WAttr("afterAutospacing")) == true)) return false;
        bool inBody = source.Parent.IsW("body");
        if (inBody && paragraph.GetBodySection() is { } section && section.Node().WChild("lnNumType").Exists) return false;
        if (properties.WChild("rPr").WChild("vanish").OnOff() == true) return false;
        foreach (OfficeXmlNode element in source.DescendantElements()) {
            if (element.NamespaceId != WordXml.W) continue;
            switch (element.LocalName) {
                case "ins": case "del": case "moveFrom": case "moveTo":
                case "rPrChange": case "pPrChange": case "numberingChange":
                case "sdt": case "customXml":
                case "rtl":
                    return false;
                case "vanish" when element.OnOff() == true:
                    return false;
                case "hyperlink" when element.Attribute(WordXml.R, "id") != null && !inBody:
                    return false;
            }
        }
        if (GetNativeRuns(paragraph).Any(run => IsNativeHiddenTextRun(run, paragraph))) return false;
        return !IsNativeParagraphInCrossParagraphField(source);
    }

    /// <summary>True when a complex field is open at the paragraph start or stays open at its end.</summary>
    private static bool IsNativeParagraphInCrossParagraphField(OfficeXmlNode paragraph) {
        if (FieldDepthChange(paragraph, out bool closesEarlier) != 0 || closesEarlier) return true;
        int depth = 0;
        foreach (OfficeXmlNode previous in paragraph.ElementsBefore()) {
            if (!previous.IsW("p")) continue;
            depth += FieldDepthChange(previous, out _);
        }
        return depth > 0;

        static int FieldDepthChange(OfficeXmlNode item, out bool closesOpenField) {
            int balance = 0;
            closesOpenField = false;
            foreach (OfficeXmlNode fieldChar in item.Descendants(WordXml.W, "fldChar")) {
                string? type = fieldChar.WAttr("fldCharType");
                if (type == "begin") balance++;
                else if (type == "end") {
                    balance--;
                    if (balance < 0) closesOpenField = true;
                }
            }
            return balance;
        }
    }

    private static bool HasNativeFrameParagraphBorders(WordParagraph paragraph) {
        NativeParagraphBorders borders = MergeNativeParagraphBorders(GetNativeParagraphStyleDefaults(paragraph).Borders,
            paragraph.Node().WChild("pPr").WChild("pBdr"));
        static bool Visible(NativeParagraphBorderSide side) => !side.IsEmpty &&
            side.Style != "none" && side.Style != "nil";
        return Visible(borders.Top) || Visible(borders.Bottom) || Visible(borders.Left) || Visible(borders.Right) || Visible(borders.Between);
    }

    /// <summary>Returns the effective frame of a paragraph: direct formatting first, then its style chain.</summary>
    private static OfficeXmlNode ResolveNativeParagraphFrame(WordParagraph paragraph) {
        OfficeXmlNode source = paragraph.Node();
        if (source.IsNone) return default;
        OfficeXmlNode frame = source.WChild("pPr").WChild("framePr");
        if (frame.IsNone) {
            foreach (var style in GetNativeParagraphStyleChain(paragraph._document, paragraph.StyleId).Reverse()) {
                frame = style.WChild("pPr").WChild("framePr");
                if (frame.Exists) break;
            }
        }

        return frame;
    }

    private static bool IsNativeDropCapFrame(OfficeXmlNode frame, string? location = null) =>
        frame.WAttr("dropCap") is { } dropCap && dropCap != "none" &&
        (location == null || dropCap == location);

    private static WordTable CreateNativeFrameTable(WordDocument document, OfficeXmlNode frame,
        IReadOnlyList<WordParagraph> paragraphs, bool notBeside) {
        int? width = ParseNativeFrameTwips(frame.WAttr("w"));
        OfficeXmlNode table = CreateNativeDetachedBodyElement(document, "tbl");
        // A reserved, undefined style keeps the document's default table style from formatting
        // frame text that was never in a table.
        OfficeXmlNode tableProperties = AppendNativeFrameElement(table, "tblPr");
        AppendNativeFrameElement(tableProperties, "tblStyle", ("val", NativeFrameTableStyleId));
        if (!notBeside) {
            string? hSpace = FormatNativeFrameTwips(ParseNativeFrameTwips(frame.WAttr("hSpace")));
            string? vSpace = FormatNativeFrameTwips(ParseNativeFrameTwips(frame.WAttr("vSpace")));
            string? vAnchor = frame.WAttr("vAnchor");
            string? hAnchor = frame.WAttr("hAnchor");
            OfficeXmlNode position = AppendNativeFrameElement(tableProperties, "tblpPr",
                ("leftFromText", hSpace), ("rightFromText", hSpace), ("topFromText", vSpace), ("bottomFromText", vSpace),
                ("vertAnchor", vAnchor == "text" ? "text" : vAnchor == "margin" ? "margin" : "page"),
                ("horzAnchor", hAnchor == "text" ? "text" : hAnchor == "margin" ? "margin" : "page"));
            // A margin drop cap hangs against the text area, in the page margin.
            if (IsNativeDropCapFrame(frame, "margin")) {
                position.SetAttribute(WordXml.W, "horzAnchor", "page");
                position.SetAttribute(WordXml.W, "tblpXSpec", "left");
            } else if (frame.WAttr("xAlign") is { } xAlign) position.SetAttribute(WordXml.W, "tblpXSpec", xAlign);
            else position.SetAttribute(WordXml.W, "tblpX", FormatNativeFrameTwips(ParseNativeFrameTwips(frame.WAttr("x")) ?? 0));
            if (frame.WAttr("yAlign") is { } yAlign && yAlign != "inline") position.SetAttribute(WordXml.W, "tblpYSpec", yAlign);
            else position.SetAttribute(WordXml.W, "tblpY", FormatNativeFrameTwips(ParseNativeFrameTwips(frame.WAttr("y")) ?? 0));
            AppendNativeFrameElement(tableProperties, "tblOverlap", ("val", "overlap"));
        } else {
            ApplyNativeNotBesideFramePlacement(tableProperties, frame);
        }

        AppendNativeFrameElement(tableProperties, "tblW", width.HasValue
            ? new[] { ("w", FormatNativeFrameTwips(width)), ("type", (string?)"dxa") }
            : new[] { ("w", (string?)"0"), ("type", (string?)"auto") });
        OfficeXmlNode borders = AppendNativeFrameElement(tableProperties, "tblBorders");
        foreach (string side in new[] { "top", "left", "bottom", "right", "insideH", "insideV" })
            AppendNativeFrameElement(borders, side, ("val", "none"));
        AppendNativeFrameElement(tableProperties, "tblLayout", ("type", width.HasValue ? "fixed" : "autofit"));
        OfficeXmlNode margins = AppendNativeFrameElement(tableProperties, "tblCellMar");
        foreach (string side in new[] { "top", "left", "bottom", "right" })
            AppendNativeFrameElement(margins, side, ("w", "0"), ("type", "dxa"));

        OfficeXmlNode grid = AppendNativeFrameElement(table, "tblGrid");
        AppendNativeFrameElement(grid, "gridCol", ("w", FormatNativeFrameTwips(width ?? 0)));
        OfficeXmlNode row = AppendNativeFrameElement(table, "tr");
        uint? height = WordXml.ParseUInt32(frame.WAttr("h"));
        if (height is uint frameHeight && frameHeight > 0) {
            OfficeXmlNode rowProperties = AppendNativeFrameElement(row, "trPr");
            AppendNativeFrameElement(rowProperties, "trHeight",
                ("val", Math.Min(frameHeight, 31680U).ToString(System.Globalization.CultureInfo.InvariantCulture)),
                ("hRule", frame.WAttr("hRule") == "exact" ? "exact" : "atLeast"));
        }
        OfficeXmlNode cell = AppendNativeFrameElement(row, "tc");
        OfficeXmlNode cellProperties = AppendNativeFrameElement(cell, "tcPr");
        AppendNativeFrameElement(cellProperties, "tcW", width.HasValue
            ? new[] { ("w", FormatNativeFrameTwips(width)), ("type", (string?)"dxa") }
            : new[] { ("w", (string?)"0"), ("type", (string?)"auto") });
        for (int i = 0; i < paragraphs.Count; i++) {
            OfficeXmlNode clone = cell.Append(paragraphs[i].Node().Clone());
            OfficeXmlNode properties = clone.GetOrPrependChild(WordXml.W, "pPr");
            OfficeXmlNode copiedFrame = properties.WChild("framePr");
            if (copiedFrame.Exists) copiedFrame.Remove();
            // A floating frame ends at its last line: the final space after is outside it.
            if (i == paragraphs.Count - 1 && !notBeside) {
                OfficeXmlNode spacing = properties.WChild("spacing");
                if (spacing.IsNone) spacing = properties.InsertOrdered(properties.CreateElement(WordXml.W, "spacing"), WPPrView.Order);
                spacing.SetAttribute(WordXml.W, "after", "0");
                spacing.SetAttribute(WordXml.W, "afterLines", null);
                spacing.SetAttribute(WordXml.W, "afterAutospacing", null);
            }
        }

        return WordConverterNodes.CreateTable(document, table);
    }

    /// <summary>Appends a new <c>w:</c> child with these attributes (a null value is left out) and returns it.</summary>
    private static OfficeXmlNode AppendNativeFrameElement(OfficeXmlNode parent, string localName, params (string Name, string? Value)[] attributes) {
        OfficeXmlNode element = parent.Append(parent.CreateElement(WordXml.W, localName));
        foreach ((string name, string? value) in attributes) {
            if (value != null) element.SetAttribute(WordXml.W, name, value);
        }
        return element;
    }

    private static string? FormatNativeFrameTwips(int? twips) =>
        twips?.ToString(System.Globalization.CultureInfo.InvariantCulture);

    private static PdfCore.PdfTablePosition CreateNativeFramePosition(PdfCore.PdfTablePosition position, OfficeXmlNode frame) =>
        // Word treats a line that starts exactly at the frame's bottom edge as beside the
        // frame, so the clearance extends by one twip. Frames without wrapping let text
        // run underneath them. A margin drop cap keeps its hSpace clear of the text by
        // moving further into the margin.
        new(position.HorizontalAnchor, position.VerticalAnchor, position.HorizontalAlignment, position.VerticalAlignment,
            IsNativeDropCapFrame(frame, "margin")
                ? global::OfficeIMO.OfficeLength.Points(position.HorizontalOffset.ToPoints() - position.DistanceRight.ToPoints())
                : position.HorizontalOffset, position.VerticalOffset, position.DistanceLeft, position.DistanceRight, position.DistanceTop,
            global::OfficeIMO.OfficeLength.Points(position.DistanceBottom.ToPoints() + 0.05D),
            position.MirrorHorizontalOnEvenPages, position.AllowOverlap) {
            WrapsText = frame.WAttr("wrap") != "none",
            AlignsInPageMargin = frame.WAttr("xAlign") is { } xAlign && xAlign != "center" ||
                IsNativeDropCapFrame(frame, "margin")
        };

    /// <summary>A notBeside frame keeps its own lines in the flow; only its horizontal placement applies.</summary>
    private static void ApplyNativeNotBesideFramePlacement(OfficeXmlNode properties, OfficeXmlNode frame) {
        string? xAlign = frame.WAttr("xAlign");
        if (xAlign == "right" || xAlign == "outside") {
            AppendNativeFrameElement(properties, "jc", ("val", "right"));
        } else if (xAlign == "center") {
            AppendNativeFrameElement(properties, "jc", ("val", "center"));
        } else if (xAlign == null && ParseNativeFrameTwips(frame.WAttr("x")) is int x && x != 0 &&
                   frame.WAttr("hAnchor") != "page") {
            AppendNativeFrameElement(properties, "tblInd", ("w", FormatNativeFrameTwips(x)), ("type", "dxa"));
        }
    }

    private const string NativeFrameTableStyleId = "OfficeIMONativeFrameProjection";

    private static int? ParseNativeFrameTwips(string? value) =>
        int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int twips)
            ? Math.Max(-31680, Math.Min(31680, twips)) : null;
}
