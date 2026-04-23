namespace OfficeIMO.Markdown;

/// <summary>
/// Inline parsing helpers for <see cref="MarkdownReader"/>.
/// </summary>
public static partial class MarkdownReader {
    private static InlineSequence ParseInlines(string text, MarkdownReaderOptions options, MarkdownReaderState? state = null, MarkdownInlineSourceMap? sourceMap = null) {
        var sequence = ParseInlinesInternal(text, options, state, allowLinks: true, allowImages: true, sourceMap);
        NormalizeInlineSequenceInPlace(sequence, options.InputNormalization);
        return sequence;
    }

    private static InlineSequence ParseInlinesInternal(string text, MarkdownReaderOptions options, MarkdownReaderState? state, bool allowLinks, bool allowImages, MarkdownInlineSourceMap? sourceMap = null) {
        var root = new InlineSequence { AutoSpacing = false };
        if (string.IsNullOrEmpty(text)) return root;
        var inlineParserExtensions = BuildEffectiveInlineParserExtensions(options);

        // We parse emphasis/strong/strikethrough using a simple stack of open frames so that nesting like
        // "*a **b** c*" behaves intuitively. This is not a full spec implementation, but it's materially
        // more robust than naive IndexOf-based matching.
        var stack = new Stack<InlineFrame>();
        stack.Push(new InlineFrame(FrameKind.Root, '\0', 0, root, -1));
        InlineSequence Current() => stack.Peek().Seq;
        MarkdownInlineSourceMap? SliceMap(int start, int length) => sourceMap?.Slice(start, length);
        void AddRawNode(IMarkdownInline node, int start, int length) {
            MarkdownInlineSourceSpans.Set(node, sourceMap?.GetSpan(start, length));
            Current().AddRaw(node);
        }
        void AddTextNode(string literal, int start, int length) => AddRawNode(new TextRun(literal), start, length);
        void AddHardBreakNode(int start, int length) => AddRawNode(new HardBreakInline(), start, length);
        InlineSequence ParseNestedInlineSegment(int relativeStart, int length, bool nestedAllowLinks, bool nestedAllowImages) {
            if (relativeStart < 0 || length <= 0 || relativeStart >= text.Length) {
                return new InlineSequence();
            }

            var safeLength = Math.Min(length, text.Length - relativeStart);
            if (safeLength <= 0) {
                return new InlineSequence();
            }

            return ParseInlinesInternal(
                text.Substring(relativeStart, safeLength),
                options,
                state,
                nestedAllowLinks,
                nestedAllowImages,
                SliceMap(relativeStart, safeLength));
        }

        int pos = 0;
        while (pos < text.Length) {
            // Hard break signal encoded by paragraph joiner as a bare '\n'
            if (text[pos] == '\n') { AddHardBreakNode(pos, 1); pos++; continue; }
            // HTML-style line breaks in source (commonly used inside table cells): <br>, <br/>, <br />
            if (options.InlineHtml && text[pos] == '<') {
                const string br = "<br>";
                const string brSelf = "<br/>";
                const string brSelfSpaced = "<br />";
                if (text.Length - pos >= br.Length && string.Compare(text, pos, br, 0, br.Length, StringComparison.OrdinalIgnoreCase) == 0) {
                    AddHardBreakNode(pos, br.Length); pos += br.Length; continue;
                }
                if (text.Length - pos >= brSelf.Length && string.Compare(text, pos, brSelf, 0, brSelf.Length, StringComparison.OrdinalIgnoreCase) == 0) {
                    AddHardBreakNode(pos, brSelf.Length); pos += brSelf.Length; continue;
                }
                if (text.Length - pos >= brSelfSpaced.Length && string.Compare(text, pos, brSelfSpaced, 0, brSelfSpaced.Length, StringComparison.OrdinalIgnoreCase) == 0) {
                    AddHardBreakNode(pos, brSelfSpaced.Length); pos += brSelfSpaced.Length; continue;
                }
            }

            if (TryParseInlineExtension(
                text,
                pos,
                options,
                state,
                allowLinks,
                allowImages,
                sourceMap,
                inlineParserExtensions,
                ParseNestedInlineSegment,
                out var extensionResult)) {
                AddRawNode(extensionResult.Inline, pos, extensionResult.ConsumedLength);
                pos += extensionResult.ConsumedLength;
                continue;
            }

            // Backslash escape (CommonMark-ish): only escape punctuation we care about so that Windows paths like
            // "C:\Support\GitHub" keep their backslashes.
            if (text[pos] == '\\') {
                if (pos + 1 < text.Length) {
                    char next = text[pos + 1];
                    if (IsBackslashEscapable(next)) {
                        AddTextNode(next.ToString(), pos, 2);
                        pos += 2;
                        continue;
                    }
                }
                AddTextNode("\\", pos, 1);
                pos++;
                continue;
            }

            // Autolink: http(s)://... until whitespace or closing punct
            if (options.AutolinkUrls && StartsWithHttp(text, pos, out int urlEnd)) {
                var url = text.Substring(pos, urlEnd - pos);
                var resolved = ResolveUrl(url, options);
                if (resolved is null) {
                    AddTextNode(url, pos, urlEnd - pos);
                } else {
                    AddRawNode(new LinkInline(url, resolved!, null), pos, urlEnd - pos);
                }
                pos = urlEnd; continue;
            }

            // Autolink: www.example.com
            if (options.AutolinkWwwUrls && StartsWithWww(text, pos, out int wwwEnd)) {
                var label = text.Substring(pos, wwwEnd - pos);
                var scheme = string.IsNullOrWhiteSpace(options.AutolinkWwwScheme) ? "https://" : options.AutolinkWwwScheme.Trim();
                if (!scheme.EndsWith("://", StringComparison.Ordinal)) scheme = scheme.TrimEnd('/') + "://";
                var href = scheme + label;
                var resolved = ResolveUrl(href, options);
                if (resolved is null) {
                    AddTextNode(label, pos, wwwEnd - pos);
                } else {
                    AddRawNode(new LinkInline(label, resolved!, null), pos, wwwEnd - pos);
                }
                pos = wwwEnd; continue;
            }

            // Autolink: plain email
            if (options.AutolinkEmails && TryConsumePlainEmail(text, pos, out int emailEnd, out string email)) {
                var href = "mailto:" + email;
                var resolved = ResolveUrl(href, options);
                if (resolved is null) {
                    AddTextNode(email, pos, emailEnd - pos);
                } else {
                    AddRawNode(new LinkInline(email, resolved!, null), pos, emailEnd - pos);
                }
                pos = emailEnd; continue;
            }
            if (text[pos] == '`') {
                // Support multi-backtick code spans: count fence length and find a matching run
                int fenceLen = 0; int k = pos; while (k < text.Length && text[k] == '`') { fenceLen++; k++; }
                int j = k; int matchStart = -1;
                while (j < text.Length) {
                    if (text[j] != '`') {
                        j++;
                        continue;
                    }

                    int candidateStart = j;
                    int candidateLen = 0;
                    while (j < text.Length && text[j] == '`') {
                        candidateLen++;
                        j++;
                    }

                    if (candidateLen == fenceLen) {
                        matchStart = candidateStart;
                        break;
                    }
                }
                if (matchStart >= 0) {
                    int contentStart = pos + fenceLen;
                    int contentLen = matchStart - contentStart;
                    if (contentLen < 0) contentLen = 0;
                    var inner = text.Substring(contentStart, contentLen);
                    inner = NormalizeCodeSpanContent(inner);
                    AddRawNode(new CodeSpanInline(inner), pos, matchStart + fenceLen - pos);
                    pos = matchStart + fenceLen; continue;
                }

                if (fenceLen > 1) {
                    AddTextNode(new string('`', fenceLen), pos, fenceLen);
                    pos += fenceLen;
                    continue;
                }
            }

            // Footnote ref [^id] should be recognized before generic link parsing
            if (options.Footnotes && text[pos] == '[' && pos + 2 < text.Length && text[pos + 1] == '^') {
                int rb = text.IndexOf(']', pos + 2);
                if (rb > pos + 2) { var lab = text.Substring(pos + 2, rb - (pos + 2)); AddRawNode(new FootnoteRefInline(lab), pos, rb + 1 - pos); pos = rb + 1; continue; }
            }

            if (TryParseImageLink(
                text,
                pos,
                out int consumed,
                out var alt2,
                out var img2,
                out var imgTitle2,
                out var href2,
                out var hrefTitle2,
                out int altStart2,
                out int altLength2,
                out int imgStart2,
                out int imgLength2,
                out int? imgTitleStart2,
                out int? imgTitleLength2,
                out int imageLinkHrefStart,
                out int imageLinkHrefLength,
                out int? imageLinkHrefTitleStart,
                out int? imageLinkHrefTitleLength)) {
                if (allowLinks && allowImages) {
                    var imgResolved = ResolveUrl(img2, options);
                    var hrefResolved = ResolveUrl(href2, options);
                    if (imgResolved is null || hrefResolved is null) {
                        // Unsafe URLs: keep content as plain text instead of a clickable linked image.
                        AddTextNode(string.IsNullOrEmpty(alt2) ? "image" : ExtractImageAltPlainText(alt2, options, state), pos, consumed);
                    } else {
                        var plainAlt2 = ExtractImageAltPlainText(alt2, options, state);
                        var imageLink = new ImageLinkInline(alt2, imgResolved!, hrefResolved!, imgTitle2, hrefTitle2, plainAlt2);
                        AddRawNode(imageLink, pos, consumed);
                        MarkdownInlineMetadataSourceSpans.SetImageLinkParts(
                            imageLink,
                            sourceMap?.GetSpan(altStart2, altLength2),
                            sourceMap?.GetSpan(imgStart2, imgLength2),
                            imgTitleStart2.HasValue && imgTitleLength2.HasValue
                                ? sourceMap?.GetSpan(imgTitleStart2.Value, imgTitleLength2.Value)
                                : null,
                            sourceMap?.GetSpan(imageLinkHrefStart, imageLinkHrefLength),
                            imageLinkHrefTitleStart.HasValue && imageLinkHrefTitleLength.HasValue
                                ? sourceMap?.GetSpan(imageLinkHrefTitleStart.Value, imageLinkHrefTitleLength.Value)
                                : null);
                    }
                    pos += consumed; continue;
                }
            }

            if (text[pos] == '!') {
                if (allowImages) {
                    if (options.Footnotes
                        && pos + 3 < text.Length
                        && text[pos + 1] == '['
                        && text[pos + 2] == '^') {
                        AddTextNode("!", pos, 1);
                        pos++;
                        continue;
                    }

                    // Reference-style image: ![alt][label], ![alt][], or shortcut ![label]
                    if (state != null && TryParseReferenceImage(
                        text,
                        pos,
                        out int consumedRefImg,
                        out var altRef,
                        out var refLabel,
                        out int altRefStart,
                        out int altRefLength)) {
                        var key = NormalizeReferenceLabel(refLabel);
                        if (state.LinkRefs.TryGetValue(key, out var defImg)) {
                            var resolved = ResolveUrl(defImg.Url, options);
                            if (resolved is null) {
                                AddTextNode(string.IsNullOrEmpty(altRef) ? "image" : ExtractImageAltPlainText(altRef, options, state), pos, consumedRefImg);
                            } else {
                                var plainAltRef = ExtractImageAltPlainText(altRef, options, state);
                                var image = new ImageInline(altRef, resolved!, defImg.Title, plainAltRef);
                                AddRawNode(image, pos, consumedRefImg);
                                MarkdownInlineMetadataSourceSpans.SetImageParts(
                                    image,
                                    sourceMap?.GetSpan(altRefStart, altRefLength),
                                    defImg.UrlSourceSpan,
                                    defImg.TitleSourceSpan);
                            }
                        } else {
                            // Preserve literal syntax when the definition is missing.
                            AddTextNode(text.Substring(pos, consumedRefImg), pos, consumedRefImg);
                        }
                        pos += consumedRefImg; continue;
                    }

                    // Inline image: ![alt](src "title")
                    if (TryParseInlineImage(
                        text,
                        pos,
                        out int consumedImg,
                        out var altImg,
                        out var srcImg,
                        out var titleImg,
                        out int altStartImg,
                        out int altLengthImg,
                        out int srcStartImg,
                        out int srcLengthImg,
                        out int? titleStartImg,
                        out int? titleLengthImg)) {
                        var srcResolved = ResolveUrl(srcImg, options);
                        if (srcResolved is null) {
                            AddTextNode(string.IsNullOrEmpty(altImg) ? "image" : ExtractImageAltPlainText(altImg, options, state), pos, consumedImg);
                        } else {
                            var plainAltImg = ExtractImageAltPlainText(altImg, options, state);
                            var image = new ImageInline(altImg, srcResolved!, titleImg, plainAltImg);
                            AddRawNode(image, pos, consumedImg);
                            MarkdownInlineMetadataSourceSpans.SetImageParts(
                                image,
                                sourceMap?.GetSpan(altStartImg, altLengthImg),
                                sourceMap?.GetSpan(srcStartImg, srcLengthImg),
                                titleStartImg.HasValue && titleLengthImg.HasValue
                                    ? sourceMap?.GetSpan(titleStartImg.Value, titleLengthImg.Value)
                                    : null);
                        }
                        pos += consumedImg; continue;
                    }

                    if (TryConsumeLiteralInlineImage(text, pos, out int literalImageLength)) {
                        AddTextNode(text.Substring(pos, literalImageLength), pos, literalImageLength);
                        pos += literalImageLength; continue;
                    }
                }

                // If image parsing does not match, keep the punctuation literal and let the next
                // iteration re-evaluate the following '[' token as a link or footnote reference.
                AddTextNode("!", pos, 1);
                pos++;
                continue;
            }

            // Angle-bracket autolinks: <https://example.com>, <mailto:user@example.com>, <tel:+123>, <user@example.com>
            if (text[pos] == '<' && TryParseAngleAutolink(text, pos, out int consumedAngle, out var labelAngle, out var hrefAngle)) {
                var resolved = ResolveUrl(hrefAngle, options);
                if (resolved is null) {
                    AddTextNode(text.Substring(pos, consumedAngle), pos, consumedAngle);
                } else {
                    var link = new LinkInline(labelAngle, resolved!, null);
                    AddRawNode(link, pos, consumedAngle);
                    MarkdownInlineMetadataSourceSpans.SetLinkParts(
                        link,
                        sourceMap?.GetSpan(pos + 1, Math.Max(0, consumedAngle - 2)),
                        null);
                }
                pos += consumedAngle;
                continue;
            }
            if (text[pos] == '[') {
                if (allowLinks) {
                    if (TryParseLink(text, pos, out int consumed2, out var label2, out var href3, out var title2, out int hrefStart2, out int hrefLength2, out int? titleStart2, out int? titleLength2)) {
                        var labelSeq = ParseInlinesInternal(label2, options, state, allowLinks: false, allowImages: false, SliceMap(pos + 1, label2.Length));

                        // Allow empty href: commonly used as placeholder or to be filled by the host.
                        if (string.IsNullOrWhiteSpace(href3)) {
                            var link = new LinkInline(labelSeq, string.Empty, title2);
                            AddRawNode(link, pos, consumed2);
                            MarkdownInlineMetadataSourceSpans.SetLinkParts(
                                link,
                                hrefLength2 > 0 ? sourceMap?.GetSpan(hrefStart2, hrefLength2) : null,
                                titleStart2.HasValue && titleLength2.HasValue ? sourceMap?.GetSpan(titleStart2.Value, titleLength2.Value) : null);
                        } else {
                            var hrefResolved = ResolveUrl(href3, options);
                            if (hrefResolved is null) {
                                // Unsafe URLs: keep the label as plain inline content instead of producing an <a href="...">.
                                foreach (var n in labelSeq.Nodes) Current().AddRaw(n);
                            } else {
                                var link = new LinkInline(labelSeq, hrefResolved!, title2);
                                AddRawNode(link, pos, consumed2);
                                MarkdownInlineMetadataSourceSpans.SetLinkParts(
                                    link,
                                    hrefLength2 > 0 ? sourceMap?.GetSpan(hrefStart2, hrefLength2) : null,
                                    titleStart2.HasValue && titleLength2.HasValue ? sourceMap?.GetSpan(titleStart2.Value, titleLength2.Value) : null);
                            }
                        }
                        pos += consumed2; continue;
                    }

                    if (state != null && TryParseCollapsedRef(text, pos, out int consumedC, out var lbl2)) {
                        var key = NormalizeReferenceLabel(lbl2);
                        var labelSeq = ParseInlinesInternal(lbl2, options, state, allowLinks: false, allowImages: false, SliceMap(pos + 1, lbl2.Length));
                        if (state.LinkRefs.TryGetValue(key, out var def2)) {
                            var resolved = ResolveUrl(def2.Url, options);
                            if (resolved is null) {
                                foreach (var n in labelSeq.Nodes) Current().AddRaw(n);
                            } else {
                                var link = new LinkInline(labelSeq, resolved!, def2.Title);
                                AddRawNode(link, pos, consumedC);
                                MarkdownInlineMetadataSourceSpans.SetLinkParts(link, def2.UrlSourceSpan, def2.TitleSourceSpan);
                            }
                            pos += consumedC; continue;
                        }

                        AddTextNode(text.Substring(pos, consumedC), pos, consumedC);
                        pos += consumedC; continue;
                    }
                    if (state != null && TryParseRefLink(text, pos, out int consumedR, out var lbl, out var refLabel)) {
                        var key = NormalizeReferenceLabel(refLabel);
                        var labelSeq = ParseInlinesInternal(lbl, options, state, allowLinks: false, allowImages: false, SliceMap(pos + 1, lbl.Length));
                        if (state.LinkRefs.TryGetValue(key, out var def)) {
                            var resolved = ResolveUrl(def.Url, options);
                            if (resolved is null) {
                                foreach (var n in labelSeq.Nodes) Current().AddRaw(n);
                            } else {
                                var link = new LinkInline(labelSeq, resolved!, def.Title);
                                AddRawNode(link, pos, consumedR);
                                MarkdownInlineMetadataSourceSpans.SetLinkParts(link, def.UrlSourceSpan, def.TitleSourceSpan);
                            }
                            pos += consumedR; continue;
                        }
                    }
                    if (state != null && TryParseShortcutRef(text, pos, out int consumedS, out var lbl3)) {
                        var key = NormalizeReferenceLabel(lbl3);
                        var labelSeq = ParseInlinesInternal(lbl3, options, state, allowLinks: false, allowImages: false, SliceMap(pos + 1, lbl3.Length));
                        if (state.LinkRefs.TryGetValue(key, out var def3)) {
                            var resolved = ResolveUrl(def3.Url, options);
                            if (resolved is null) {
                                foreach (var n in labelSeq.Nodes) Current().AddRaw(n);
                            } else {
                                var link = new LinkInline(labelSeq, resolved!, def3.Title);
                                AddRawNode(link, pos, consumedS);
                                MarkdownInlineMetadataSourceSpans.SetLinkParts(link, def3.UrlSourceSpan, def3.TitleSourceSpan);
                            }
                            pos += consumedS; continue;
                        }

                        AddTextNode(text.Substring(pos, consumedS), pos, consumedS);
                        pos += consumedS; continue;
                    }
                }
            }

            // Emphasis / strong / strike / highlight using delimiter-run rules + an open-frame stack.
            if (text[pos] == '*' || text[pos] == '_' || text[pos] == '~' || text[pos] == '=') {
                char marker = text[pos];
                int runLen = 1;
                while (pos + runLen < text.Length && text[pos + runLen] == marker) runLen++;

                bool splitDoubleRunIntoDualItalic = ShouldSplitDoubleRunIntoDualItalic(text, pos, marker, runLen, stack);

                if (ShouldTreatDelimiterRunAsLiteral(text, pos, marker, runLen, stack, splitDoubleRunIntoDualItalic, out int literalRunLength)) {
                    AddTextNode(new string(marker, literalRunLength), pos, literalRunLength);
                    pos += literalRunLength;
                    continue;
                }

                if (ShouldTreatSingleMarkerAsLiteralInsideBold(text, pos, marker, runLen, stack)) {
                    AddTextNode(marker.ToString(), pos, 1);
                    pos++;
                    continue;
                }

                // "==" always requires a double delimiter. "~" can opt into cmark-gfm style single-tilde strike.
                if (marker == '=' && runLen < 2) {
                    AddTextNode(marker.ToString(), pos, 1);
                    pos++;
                    continue;
                }

                if (marker == '~' && runLen < (options.SingleTildeStrikethrough ? 1 : 2)) {
                    AddTextNode(marker.ToString(), pos, 1);
                    pos++;
                    continue;
                }

                GetDelimiterFlags(text, pos, marker, runLen, out bool canOpen, out bool canClose);

                if (ShouldTreatMixedSingleMarkerAsLiteral(text, pos, marker, runLen, canOpen, canClose, stack)) {
                    AddTextNode(marker.ToString(), pos, 1);
                    pos++;
                    continue;
                }

                int remaining = runLen;
                if (canClose) {
                    while (remaining > 0) {
                        if (TryRebalanceParentBoldWithInnerItalicIntoDualItalic(text, pos, stack, marker, remaining, out int dualItalicRebalanced)) {
                            remaining -= dualItalicRebalanced;
                            continue;
                        }

                        if (!TryRebalanceLeadingBoldInsideItalic(stack, marker, remaining, out int rebalanced)) break;
                        remaining -= rebalanced;
                    }
                }

                bool preferInnerBold = ShouldPreferInnerBold(stack, marker, remaining, canOpen, canClose);
                bool splitDoubleUnderscoreOpener = ShouldSplitDoubleUnderscoreToLiteralAndItalic(text, pos, remaining, canOpen, canClose);
                int literalPrefixForOddCloser = GetLiteralPrefixLengthForOddCloser(text, pos, marker, remaining, canOpen, canClose);

                if (canClose && !preferInnerBold) {
                    while (remaining > 0) {
                        if (!TryCloseFrame(stack, marker, remaining, out int consumedClose)) break;
                        remaining -= consumedClose;
                    }
                }

                if (canOpen) {
                    if (splitDoubleRunIntoDualItalic) {
                        stack.Push(new InlineFrame(FrameKind.Italic, marker, 1, new InlineSequence { AutoSpacing = false }, pos));
                        stack.Push(new InlineFrame(FrameKind.Italic, marker, 1, new InlineSequence { AutoSpacing = false }, pos + 1));
                        remaining -= 2;
                    }
                    else if (preferInnerBold) {
                        stack.Push(new InlineFrame(FrameKind.Bold, marker, 2, new InlineSequence { AutoSpacing = false }, pos));
                        remaining -= 2;
                    }

                    if (splitDoubleUnderscoreOpener) {
                        AddTextNode("_", pos, 1);
                        stack.Push(new InlineFrame(FrameKind.Italic, marker, 1, new InlineSequence { AutoSpacing = false }, pos + 1));
                        remaining -= 2;
                    }
                    else if (literalPrefixForOddCloser > 0) {
                        AddTextNode(new string(marker, literalPrefixForOddCloser), pos, literalPrefixForOddCloser);
                        remaining -= literalPrefixForOddCloser;
                    }

                    while (remaining > 0) {
                        if (marker == '~') {
                            int strikeDelimiterLength = options.SingleTildeStrikethrough ? 1 : 2;
                            if (remaining >= strikeDelimiterLength) {
                                stack.Push(new InlineFrame(FrameKind.Strike, marker, strikeDelimiterLength, new InlineSequence { AutoSpacing = false }, pos + (runLen - remaining)));
                                remaining -= strikeDelimiterLength;
                                continue;
                            }
                            break;
                        }

                        if (marker == '=') {
                            if (remaining >= 2) {
                                stack.Push(new InlineFrame(FrameKind.Highlight, marker, 2, new InlineSequence { AutoSpacing = false }, pos + (runLen - remaining)));
                                remaining -= 2;
                                continue;
                            }
                            break;
                        }

                        if (remaining >= 2) {
                            stack.Push(new InlineFrame(FrameKind.Bold, marker, 2, new InlineSequence { AutoSpacing = false }, pos + (runLen - remaining)));
                            remaining -= 2;
                            continue;
                        }

                        stack.Push(new InlineFrame(FrameKind.Italic, marker, 1, new InlineSequence { AutoSpacing = false }, pos + (runLen - remaining)));
                        remaining -= 1;
                    }
                }

                if (remaining > 0) {
                    AddTextNode(new string(marker, remaining), pos + (runLen - remaining), remaining);
                }

                pos += runLen;
                continue;
            }

            if (options.InlineHtml && text[pos] == '<') {
                if (TryParseSupportedInlineHtmlTag(text, pos, options, state, allowLinks, allowImages, out int consumedHtmlTag, out var htmlNode)) {
                    AddRawNode(htmlNode, pos, consumedHtmlTag);
                    pos += consumedHtmlTag;
                    continue;
                }
            }

            // Footnote ref [^id]
            if (options.Footnotes && text[pos] == '[' && pos + 2 < text.Length && text[pos + 1] == '^') {
                int rb = text.IndexOf(']', pos + 2);
                if (rb > pos + 2) { var lab = text.Substring(pos + 2, rb - (pos + 2)); AddRawNode(new FootnoteRefInline(lab), pos, rb + 1 - pos); pos = rb + 1; continue; }
            }

            int start = pos; pos++;
            while (pos < text.Length && !IsPotentialInlineStart(text[pos], options.InlineHtml, allowLinks, allowImages)) {
                // Ensure our explicit inline handlers see these characters.
                if (text[pos] == '\n') break;
                if (text[pos] == '\\' && pos + 1 < text.Length && IsBackslashEscapable(text[pos + 1])) break;
                if (text[pos] == '<' && IsAngleAutolinkStart(text, pos)) break;
                if (options.AutolinkUrls && (text[pos] == 'h' || text[pos] == 'H') && StartsWithHttp(text, pos, out _)) break;
                if (options.AutolinkWwwUrls && (text[pos] == 'w' || text[pos] == 'W') && StartsWithWww(text, pos, out _)) break;
                if (options.AutolinkEmails && IsEmailStartChar(text[pos]) && TryConsumePlainEmail(text, pos, out _, out _)) break;
                if (inlineParserExtensions.Count > 0
                    && TryParseInlineExtension(
                        text,
                        pos,
                        options,
                        state,
                        allowLinks,
                        allowImages,
                        sourceMap,
                        inlineParserExtensions,
                        ParseNestedInlineSegment,
                        out _)) break;
                pos++;
            }
            AddTextNode(text.Substring(start, pos - start), start, pos - start);
        }

        // Unwind any unclosed emphasis frames: treat their markers as literal text.
        while (stack.Count > 1) {
            var f = stack.Pop();
            var parent = stack.Peek().Seq;
            var markerNode = new TextRun(new string(f.Marker, f.OpenLen));
            MarkdownInlineSourceSpans.Set(markerNode, sourceMap?.GetSpan(f.OpenIndex, f.OpenLen));
            parent.AddRaw(markerNode);
            foreach (var node in f.Seq.Nodes) parent.AddRaw(node);
        }

        return root;
    }

    private static string ExtractImageAltPlainText(string altMarkdown, MarkdownReaderOptions options, MarkdownReaderState? state) {
        if (string.IsNullOrEmpty(altMarkdown)) {
            return string.Empty;
        }

        var altSequence = ParseInlinesInternal(
            altMarkdown,
            options,
            state,
            allowLinks: true,
            allowImages: true);
        return InlinePlainText.Extract(altSequence);
    }

    private enum FrameKind {
        Root,
        Italic,
        Bold,
        Strike,
        Highlight,
    }

    private sealed class InlineFrame {
        public InlineFrame(FrameKind kind, char marker, int openLen, InlineSequence seq, int openIndex) {
            Kind = kind;
            Marker = marker;
            OpenLen = openLen;
            Seq = seq;
            OpenIndex = openIndex;
        }

        public FrameKind Kind { get; }
        public char Marker { get; }
        public int OpenLen { get; }
        public InlineSequence Seq { get; }
        public int OpenIndex { get; }
    }

    private static bool TryCloseFrame(Stack<InlineFrame> stack, char marker, int remaining, out int consumed) {
        consumed = 0;
        if (stack == null || stack.Count <= 1) return false;
        var top = stack.Peek();
        if (top.Marker != marker) return false;

        // Close the innermost matching frame only; this avoids crossing.
        if (top.Kind == FrameKind.Italic && remaining >= 1) {
            stack.Pop();
            var node = new ItalicSequenceInline(top.Seq);
            stack.Peek().Seq.AddRaw(node);
            consumed = 1;
            return true;
        }
        if (top.Kind == FrameKind.Bold && remaining >= 2) {
            stack.Pop();
            var node = new BoldSequenceInline(top.Seq);
            stack.Peek().Seq.AddRaw(node);
            consumed = 2;
            return true;
        }
        if (top.Kind == FrameKind.Strike && remaining >= top.OpenLen) {
            stack.Pop();
            var node = new StrikethroughSequenceInline(top.Seq);
            stack.Peek().Seq.AddRaw(node);
            consumed = top.OpenLen;
            return true;
        }
        if (top.Kind == FrameKind.Highlight && remaining >= 2) {
            stack.Pop();
            var node = new HighlightSequenceInline(top.Seq);
            stack.Peek().Seq.AddRaw(node);
            consumed = 2;
            return true;
        }
        return false;
    }

    private static bool ShouldTreatSingleMarkerAsLiteralInsideBold(string text, int start, char marker, int runLen, Stack<InlineFrame> stack) {
        if (runLen != 1) return false;
        if (marker != '*' && marker != '_') return false;
        if (string.IsNullOrEmpty(text) || start < 0 || start >= text.Length) return false;
        if (stack == null || stack.Count <= 1) return false;

        var top = stack.Peek();
        if (top.Kind != FrameKind.Bold || top.Marker != marker || top.OpenLen != 2) return false;

        int nextDoubleClose = FindNextClosingDelimiterRunIndex(text, start + 1, marker, requiredRunLength: 2);
        if (nextDoubleClose >= 0) {
            int trailingSingleClose = FindNextClosingDelimiterRunIndex(text, nextDoubleClose + 2, marker, requiredRunLength: 1);
            if (trailingSingleClose >= 0) return false;
        }

        int nextRun = FindNextDelimiterRunLength(text, start + 1, marker);
        return nextRun == 2;
    }

    private static bool ShouldTreatDelimiterRunAsLiteral(string text, int start, char marker, int runLen, Stack<InlineFrame> stack, bool splitDoubleRunIntoDualItalic, out int literalRunLength) {
        literalRunLength = 0;
        if (runLen != 2) return false;
        if (marker != '*' && marker != '_') return false;
        if (string.IsNullOrEmpty(text) || start < 0 || start >= text.Length) return false;
        if (stack == null || stack.Count <= 1) return false;

        var top = stack.Peek();
        if (top.Kind != FrameKind.Italic || top.Marker != marker || top.OpenLen != 1) return false;

        var frames = stack.ToArray();
        if (frames.Length >= 2) {
            var parent = frames[1];
            // Keep the leading triple-delimiter path available for rebalancing into <em><strong>... later.
            if (parent.Kind == FrameKind.Bold && parent.Marker == marker && parent.OpenLen == 2 && parent.Seq.Nodes.Count == 0) return false;

            if (parent.Kind == FrameKind.Bold && parent.Marker == marker && parent.OpenLen == 2 && parent.Seq.Nodes.Count > 0) {
                int trailingSingleClose = FindNextClosingDelimiterRunIndex(text, start + 2, marker, requiredRunLength: 1);
                if (trailingSingleClose >= 0) return false;
            }
        }

        if (splitDoubleRunIntoDualItalic) return false;

        int nextRun = FindNextDelimiterRunLength(text, start + 2, marker);
        if (nextRun != 1) return false;

        literalRunLength = 2;
        return true;
    }

    private static int FindNextDelimiterRunLength(string text, int start, char marker) {
        if (string.IsNullOrEmpty(text)) return 0;
        for (int i = Math.Max(0, start); i < text.Length; i++) {
            if (text[i] != marker) continue;

            int run = 1;
            while (i + run < text.Length && text[i + run] == marker) run++;
            return run;
        }
        return 0;
    }

    private static bool TryRebalanceLeadingBoldInsideItalic(Stack<InlineFrame> stack, char marker, int remaining, out int consumed) {
        consumed = 0;
        if (remaining < 2) return false;
        if (marker != '*' && marker != '_') return false;
        if (stack == null || stack.Count < 3) return false;

        var frames = stack.ToArray();
        var top = frames[0];
        var parent = frames[1];
        if (top.Kind != FrameKind.Italic || top.Marker != marker || top.OpenLen != 1) return false;
        if (parent.Kind != FrameKind.Bold || parent.Marker != marker || parent.OpenLen != 2) return false;
        if (parent.Seq.Nodes.Count != 0) return false;

        stack.Pop();
        stack.Pop();

        var italic = new InlineFrame(FrameKind.Italic, marker, 1, new InlineSequence { AutoSpacing = false }, parent.OpenIndex);
        italic.Seq.AddRaw(new BoldSequenceInline(top.Seq));
        stack.Push(italic);
        consumed = 2;
        return true;
    }

    private static bool TryRebalanceParentBoldWithInnerItalicIntoDualItalic(string text, int start, Stack<InlineFrame> stack, char marker, int remaining, out int consumed) {
        consumed = 0;
        if (remaining != 2) return false;
        if (marker != '*' && marker != '_') return false;
        if (string.IsNullOrEmpty(text) || start < 0 || start >= text.Length) return false;
        if (stack == null || stack.Count < 3) return false;

        var frames = stack.ToArray();
        var top = frames[0];
        var parent = frames[1];
        if (top.Kind != FrameKind.Italic || top.Marker != marker || top.OpenLen != 1) return false;
        if (parent.Kind != FrameKind.Bold || parent.Marker != marker || parent.OpenLen != 2) return false;
        if (parent.Seq.Nodes.Count == 0) return false;

        int trailingSingleClose = FindNextClosingDelimiterRunIndex(text, start + 2, marker, requiredRunLength: 1);
        if (trailingSingleClose < 0) return false;

        stack.Pop();
        stack.Pop();

        var middle = new InlineSequence { AutoSpacing = false };
        foreach (var node in parent.Seq.Nodes) {
            middle.AddRaw(node);
        }

        middle.AddRaw(new ItalicSequenceInline(top.Seq));

        var outer = new InlineFrame(FrameKind.Italic, marker, 1, new InlineSequence { AutoSpacing = false }, parent.OpenIndex);
        outer.Seq.AddRaw(new ItalicSequenceInline(middle));
        stack.Push(outer);
        consumed = 2;
        return true;
    }

    private static bool ShouldPreferInnerBold(Stack<InlineFrame> stack, char marker, int remaining, bool canOpen, bool canClose) {
        if (!canOpen || !canClose || remaining != 2) return false;
        if (marker != '*' && marker != '_') return false;
        if (stack == null || stack.Count <= 1) return false;

        var top = stack.Peek();
        return top.Marker == marker && top.Kind == FrameKind.Italic;
    }

    private static bool ShouldSplitDoubleUnderscoreToLiteralAndItalic(string text, int start, int runLen, bool canOpen, bool canClose) {
        if (!canOpen || canClose) return false;
        if (runLen != 2) return false;
        if (string.IsNullOrEmpty(text) || start < 0 || start >= text.Length) return false;
        if (text[start] != '_') return false;

        return !HasFutureClosingDelimiterRun(text, start + 2, '_', minimumRunLength: 2) &&
               HasFutureClosingDelimiterRun(text, start + 2, '_', minimumRunLength: 1);
    }

    private static int GetLiteralPrefixLengthForOddCloser(string text, int start, char marker, int runLen, bool canOpen, bool canClose) {
        if (!canOpen || canClose) return 0;
        if (runLen < 2 || (runLen % 2) != 0) return 0;
        if (marker != '*' && marker != '_') return 0;
        if (string.IsNullOrEmpty(text) || start < 0 || start >= text.Length) return 0;

        for (int candidate = runLen - 1; candidate >= 1; candidate -= 2) {
            if (FindNextClosingDelimiterRunIndex(text, start + runLen, marker, requiredRunLength: candidate) < 0) continue;

            bool hasSameOrLongerEvenCloser = false;
            for (int even = runLen; even >= candidate + 1; even -= 2) {
                if (FindNextClosingDelimiterRunIndex(text, start + runLen, marker, requiredRunLength: even) >= 0) {
                    hasSameOrLongerEvenCloser = true;
                    break;
                }
            }

            if (!hasSameOrLongerEvenCloser) {
                return runLen - candidate;
            }
        }

        return 0;
    }

    private static bool HasFutureClosingDelimiterRun(string text, int start, char marker, int minimumRunLength) {
        if (string.IsNullOrEmpty(text)) return false;
        if (minimumRunLength <= 0) minimumRunLength = 1;

        for (int i = Math.Max(0, start); i < text.Length; i++) {
            if (text[i] != marker) continue;

            int runLen = 1;
            while (i + runLen < text.Length && text[i + runLen] == marker) runLen++;

            GetDelimiterFlags(text, i, marker, runLen, out _, out bool canClose);
            if (canClose && runLen >= minimumRunLength) return true;

            i += runLen - 1;
        }

        return false;
    }

    private static bool ShouldTreatMixedSingleMarkerAsLiteral(string text, int start, char marker, int runLen, bool canOpen, bool canClose, Stack<InlineFrame> stack) {
        if (!canOpen || canClose) return false;
        if (runLen != 1) return false;
        if (marker != '*' && marker != '_') return false;
        if (string.IsNullOrEmpty(text) || start < 0 || start >= text.Length) return false;
        if (stack == null || stack.Count <= 1) return false;

        var top = stack.Peek();
        if (top.Kind != FrameKind.Italic || top.OpenLen != 1) return false;
        if (top.Marker == marker) return false;

        int outerClose = FindNextClosingDelimiterIndex(text, start + 1, top.Marker, minimumRunLength: 1);
        if (outerClose < 0) return false;

        int innerClose = FindNextClosingDelimiterIndex(text, start + 1, marker, minimumRunLength: 1);
        return innerClose < 0 || outerClose < innerClose;
    }

    private static bool ShouldSplitDoubleRunIntoDualItalic(string text, int start, char marker, int runLen, Stack<InlineFrame> stack) {
        if (runLen != 2) return false;
        if (marker != '*' && marker != '_') return false;
        if (string.IsNullOrEmpty(text) || start < 0 || start >= text.Length) return false;
        if (stack == null || stack.Count <= 1) return false;

        var top = stack.Peek();
        if (top.Kind != FrameKind.Italic || top.Marker != marker || top.OpenLen != 1) return false;

        int singleClose = FindNextClosingDelimiterRunIndex(text, start + 2, marker, requiredRunLength: 1);
        if (singleClose < 0) return false;

        int doubleClose = FindNextClosingDelimiterRunIndex(text, singleClose + 1, marker, requiredRunLength: 2);
        if (doubleClose < 0) return false;

        int afterSingle = singleClose + 1;
        return afterSingle < text.Length && char.IsWhiteSpace(text[afterSingle]);
    }

    private static int FindNextClosingDelimiterIndex(string text, int start, char marker, int minimumRunLength) {
        if (string.IsNullOrEmpty(text)) return -1;
        if (minimumRunLength <= 0) minimumRunLength = 1;

        for (int i = Math.Max(0, start); i < text.Length; i++) {
            if (text[i] != marker) continue;

            int runLen = 1;
            while (i + runLen < text.Length && text[i + runLen] == marker) runLen++;

            GetDelimiterFlags(text, i, marker, runLen, out _, out bool canClose);
            if (canClose && runLen >= minimumRunLength) return i;

            i += runLen - 1;
        }

        return -1;
    }

    private static int FindNextClosingDelimiterRunIndex(string text, int start, char marker, int requiredRunLength) {
        if (string.IsNullOrEmpty(text)) return -1;
        if (requiredRunLength <= 0) requiredRunLength = 1;

        for (int i = Math.Max(0, start); i < text.Length; i++) {
            if (text[i] != marker) continue;

            int runLen = 1;
            while (i + runLen < text.Length && text[i + runLen] == marker) runLen++;

            GetDelimiterFlags(text, i, marker, runLen, out _, out bool canClose);
            if (canClose && runLen == requiredRunLength) return i;

            i += runLen - 1;
        }

        return -1;
    }

    private static void GetDelimiterFlags(string text, int start, char marker, int runLen, out bool canOpen, out bool canClose) {
        canOpen = false;
        canClose = false;
        if (string.IsNullOrEmpty(text) || start < 0 || start >= text.Length) return;

        char prev = start > 0 ? text[start - 1] : '\0';
        int nextIndex = start + runLen;
        char next = nextIndex < text.Length ? text[nextIndex] : '\0';

        bool prevWs = prev == '\0' || char.IsWhiteSpace(prev);
        bool nextWs = next == '\0' || char.IsWhiteSpace(next);
        bool prevPunct = prev != '\0' && IsPunctuationOrSymbol(prev);
        bool nextPunct = next != '\0' && IsPunctuationOrSymbol(next);

        bool leftFlanking = !nextWs && (!nextPunct || prevWs || prevPunct);
        bool rightFlanking = !prevWs && (!prevPunct || nextWs || nextPunct);

        if (marker == '~') {
            // Pragmatic GFM-like: "~~" opens/closes when not adjacent to whitespace on the relevant side.
            canOpen = !nextWs;
            canClose = !prevWs;
            return;
        }

        if (marker == '=') {
            // Pragmatic mark/highlight handling: "==" opens/closes when it hugs non-whitespace text.
            canOpen = runLen >= 2 && !nextWs;
            canClose = runLen >= 2 && !prevWs;
            return;
        }

        if (marker == '*') {
            canOpen = leftFlanking;
            canClose = rightFlanking;
            return;
        }

        // '_' is more restrictive (avoid intraword emphasis like foo_bar_baz).
        if (marker == '_') {
            canOpen = leftFlanking && (!rightFlanking || prevPunct || prevWs);
            canClose = rightFlanking && (!leftFlanking || nextPunct || nextWs);
            return;
        }
    }

    private static bool IsPunctuationOrSymbol(char c) => char.IsPunctuation(c) || char.IsSymbol(c);

    private static string NormalizeCodeSpanContent(string inner) {
        if (inner == null) return string.Empty;

        // Normalize newlines to spaces (CommonMark-like).
        if (inner.IndexOf('\r') >= 0) inner = inner.Replace("\r\n", "\n").Replace("\r", "\n");
        if (inner.IndexOf('\n') >= 0) inner = inner.Replace("\n", " ");

        // Trim a single leading+trailing space if both exist and the content is not all spaces.
        if (inner.Length >= 2 && inner[0] == ' ' && inner[inner.Length - 1] == ' ') {
            bool anyNonSpace = false;
            for (int i = 0; i < inner.Length; i++) {
                if (inner[i] != ' ') { anyNonSpace = true; break; }
            }
            if (anyNonSpace) inner = inner.Substring(1, inner.Length - 2);
        }

        return inner;
    }

    private static bool IsAngleAutolinkStart(string text, int start) {
        if (start < 0 || start >= text.Length) return false;
        if (text[start] != '<') return false;
        return TryParseAngleAutolink(text, start, out _, out _, out _);
    }

    private static bool TryParseAngleAutolink(string text, int start, out int consumed, out string label, out string href) {
        consumed = 0;
        label = href = string.Empty;
        if (start < 0 || start >= text.Length || text[start] != '<') return false;
        int gt = text.IndexOf('>', start + 1);
        if (gt < 0) return false;
        if (gt == start + 1) return false;

        // Disallow whitespace/control inside.
        for (int i = start + 1; i < gt; i++) {
            char c = text[i];
            if (char.IsWhiteSpace(c) || char.IsControl(c)) return false;
        }

        var inner = text.Substring(start + 1, gt - (start + 1));

        // URL form
        if (inner.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
            inner.StartsWith("https://", StringComparison.OrdinalIgnoreCase)) {
            label = inner;
            href = inner;
            consumed = gt - start + 1;
            return true;
        }

        if (inner.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase)) {
            label = inner;
            href = inner;
            consumed = gt - start + 1;
            return true;
        }

        if (TryGetScheme(inner, out var scheme) && IsUriAngleAutolink(inner, scheme)) {
            label = inner;
            href = inner;
            consumed = gt - start + 1;
            return true;
        }

        // Email form
        if (LooksLikeEmail(inner)) {
            label = inner;
            href = "mailto:" + inner;
            consumed = gt - start + 1;
            return true;
        }

        return false;
    }

    private static bool IsUriAngleAutolink(string inner, string scheme) {
        if (string.IsNullOrEmpty(inner) || string.IsNullOrEmpty(scheme)) return false;

        // Match CommonMark-style absolute URI autolinks instead of limiting support to scheme://...
        // This keeps tel:, urn:, xmpp:, etc. on the same policy-controlled path as http(s)/mailto.
        if (scheme.Length < 2 || scheme.Length > 32) return false;
        if (inner.Length <= scheme.Length + 1) return false;

        for (int i = scheme.Length + 1; i < inner.Length; i++) {
            char c = inner[i];
            if (char.IsWhiteSpace(c) || char.IsControl(c) || c == '<' || c == '>') return false;
        }

        return true;
    }

    private static bool LooksLikeEmail(string s) {
        if (string.IsNullOrEmpty(s)) return false;
        int at = s.IndexOf('@');
        if (at <= 0 || at == s.Length - 1) return false;
        // Reject "mailto:" which is a URL form and will be handled above if ever enabled.
        if (s.IndexOf(':') >= 0) return false;

        string domain = s.Substring(at + 1);
        // Require at least one '.' in domain and not at the ends.
        int dot = domain.IndexOf('.');
        if (dot <= 0 || dot == domain.Length - 1) return false;

        // Basic character checks (no spaces/control already enforced by caller).
        for (int i = 0; i < s.Length; i++) {
            char c = s[i];
            if (c == '@') continue;
            if (c == '.' || c == '-' || c == '_' || c == '+') continue;
            if (char.IsLetterOrDigit(c)) continue;
            return false;
        }
        return true;
    }

    private static int FindMatchingBracket(string text, int openIndex) {
        if (string.IsNullOrEmpty(text) || openIndex < 0 || openIndex >= text.Length || text[openIndex] != '[') return -1;

        int depth = 0;
        bool escaped = false;
        for (int i = openIndex; i < text.Length; i++) {
            char c = text[i];
            if (escaped) {
                escaped = false;
                continue;
            }

            if (c == '\\') {
                escaped = true;
                continue;
            }

            if (c == '[') {
                depth++;
                continue;
            }

            if (c == ']') {
                depth--;
                if (depth == 0) return i;
            }
        }

        return -1;
    }

    private static int FindReferenceLabelEnd(string text, int openIndex) {
        if (string.IsNullOrEmpty(text) || openIndex < 0 || openIndex >= text.Length || text[openIndex] != '[') return -1;

        bool escaped = false;
        for (int i = openIndex + 1; i < text.Length; i++) {
            char c = text[i];
            if (escaped) {
                escaped = false;
                continue;
            }

            if (c == '\\') {
                escaped = true;
                continue;
            }

            if (c == '[') return -1;
            if (c == ']') return i;
        }

        return -1;
    }

    private static string UnescapeMarkdownBackslashEscapes(string value) {
        if (string.IsNullOrEmpty(value)) return value ?? string.Empty;

        var sb = new StringBuilder(value.Length);
        for (int i = 0; i < value.Length; i++) {
            char c = value[i];
            if (c == '\\' && i + 1 < value.Length && IsBackslashEscapable(value[i + 1])) {
                sb.Append(value[i + 1]);
                i++;
                continue;
            }

            sb.Append(c);
        }

        return sb.ToString();
    }

    private static bool TryParseRefLink(string text, int start, out int consumed, out string label, out string refLabel) {
        consumed = 0; label = refLabel = string.Empty;
        if (start >= text.Length || text[start] != '[') return false;
        int rb = FindMatchingBracket(text, start); if (rb < 0) return false;
        if (rb + 1 >= text.Length || text[rb + 1] != '[') return false;
        int rb2 = FindMatchingBracket(text, rb + 1); if (rb2 < 0) return false;
        label = text.Substring(start + 1, rb - (start + 1));
        refLabel = text.Substring(rb + 2, rb2 - (rb + 2));
        consumed = rb2 - start + 1; return true;
    }

    private static bool TryParseCollapsedRef(string text, int start, out int consumed, out string label) {
        consumed = 0; label = string.Empty;
        if (start >= text.Length || text[start] != '[') return false;
        int rb = FindMatchingBracket(text, start); if (rb < 0) return false;
        if (rb + 2 >= text.Length || text[rb + 1] != '[' || text[rb + 2] != ']') return false;
        label = text.Substring(start + 1, rb - (start + 1));
        consumed = rb + 3 - start;
        return true;
    }

    private static bool TryParseShortcutRef(string text, int start, out int consumed, out string label) {
        consumed = 0; label = string.Empty;
        if (start >= text.Length || text[start] != '[') return false;
        int rb = FindMatchingBracket(text, start); if (rb < 0) return false;
        if (rb + 1 < text.Length && text[rb + 1] == '[') return false;
        label = text.Substring(start + 1, rb - (start + 1));
        consumed = rb + 1 - start;
        return true;
    }

    private static bool TryConsumeLiteralInlineImage(string text, int start, out int consumed) {
        consumed = 0;
        if (start + 1 >= text.Length || text[start] != '!' || text[start + 1] != '[') return false;
        int altEnd = FindMatchingBracket(text, start + 1);
        if (altEnd < 0) return false;
        if (altEnd + 1 >= text.Length || text[altEnd + 1] != '(') return false;
        int parenClose = FindMatchingParen(text, altEnd + 1);
        if (parenClose < 0) return false;
        consumed = parenClose - start + 1;
        return true;
    }

    private static string? ResolveUrl(string url, MarkdownReaderOptions? options) {
        if (url.Length == 0) return string.Empty;
        if (string.IsNullOrWhiteSpace(url)) return null;
        url = url.Trim();

        // Block scriptable schemes by default.
        if (TryGetScheme(url, out var scheme)) {
            if (options?.RestrictUrlSchemes == true && !IsAllowedScheme(scheme, options.AllowedUrlSchemes)) return null;
            if (options?.DisallowScriptUrls != false) {
                if (scheme.Equals("javascript", StringComparison.OrdinalIgnoreCase) ||
                    scheme.Equals("vbscript", StringComparison.OrdinalIgnoreCase)) {
                    return null;
                }
            }
            if (options?.DisallowFileUrls == true) {
                if (scheme.Equals("file", StringComparison.OrdinalIgnoreCase) || IsWindowsDriveLike(url)) return null;
            }
            if (scheme.Equals("mailto", StringComparison.OrdinalIgnoreCase)) return (options?.AllowMailtoUrls ?? true) ? url : null;
            if (scheme.Equals("data", StringComparison.OrdinalIgnoreCase)) return (options?.AllowDataUrls ?? true) ? url : null;
            // http/https and unknown schemes: keep as-is (host may further restrict)
            return url;
        }

        if (url.StartsWith("//")) {
            if (options?.AllowProtocolRelativeUrls != false) {
                if (options?.RestrictUrlSchemes == true && !IsAllowedScheme("http", options.AllowedUrlSchemes) && !IsAllowedScheme("https", options.AllowedUrlSchemes)) return null;
                return url;
            }
            return null;
        }
        if (url.StartsWith("#")) return url;
        if (options?.DisallowFileUrls == true && IsWindowsDriveLike(url)) return null;

        var baseUri = options?.BaseUri;
        if (!string.IsNullOrWhiteSpace(baseUri)) {
            try {
                // Legacy behavior: only apply BaseUri when it is http(s), and only resolve into http(s).
                var baseAbs = new Uri(baseUri, UriKind.Absolute);
                if (!baseAbs.Scheme.Equals("http", StringComparison.OrdinalIgnoreCase) &&
                    !baseAbs.Scheme.Equals("https", StringComparison.OrdinalIgnoreCase)) {
                    return url;
                }
                var resolved = new Uri(baseAbs, url);
                if (!resolved.Scheme.Equals("http", StringComparison.OrdinalIgnoreCase) &&
                    !resolved.Scheme.Equals("https", StringComparison.OrdinalIgnoreCase)) {
                    return url;
                }
                return resolved.ToString();
            }
            catch (UriFormatException) { /* invalid base or relative path; keep original */ }
        }

        return url; // relative or unknown: leave as-is
    }

    private static bool IsAllowedScheme(string scheme, string[]? allowedSchemes) {
        if (string.IsNullOrEmpty(scheme)) return false;
        if (allowedSchemes == null || allowedSchemes.Length == 0) return false;
        for (int i = 0; i < allowedSchemes.Length; i++) {
            var s = allowedSchemes[i];
            if (string.IsNullOrWhiteSpace(s)) continue;
            if (scheme.Equals(s.Trim(), StringComparison.OrdinalIgnoreCase)) return true;
        }
        return false;
    }

    private static bool TryGetScheme(string url, out string scheme) {
        scheme = string.Empty;
        int colon = url.IndexOf(':');
        if (colon <= 0) return false;
        // If there's a path/query/fragment delimiter before ':', it's not a scheme.
        int slash = url.IndexOfAny(new[] { '/', '?', '#' });
        if (slash >= 0 && slash < colon) return false;
        // URI scheme must start with a letter and be [A-Za-z0-9+.-]*
        char first = url[0];
        if (!char.IsLetter(first)) return false;
        for (int i = 1; i < colon; i++) {
            char c = url[i];
            bool ok = char.IsLetterOrDigit(c) || c == '+' || c == '-' || c == '.';
            if (!ok) return false;
        }
        scheme = url.Substring(0, colon);
        return true;
    }

    private static bool IsWindowsDriveLike(string url) {
        // Treat "C:\..." and "C:/..." as file-like.
        return url.Length >= 3
               && char.IsLetter(url[0])
               && url[1] == ':'
               && (url[2] == '\\' || url[2] == '/');
    }



    private static readonly bool[] PotentialInlineStartLookup = CreatePotentialInlineStartLookup();

    private static bool[] CreatePotentialInlineStartLookup() {
        var lookup = new bool[128];
        lookup['['] = true;
        lookup['!'] = true;
        lookup['`'] = true;
        lookup['*'] = true;
        lookup['_'] = true;
        lookup['~'] = true;
        lookup['='] = true;
        return lookup;
    }

    private static bool IsBackslashEscapable(char c) {
        // CommonMark backslash-escapable punctuation (plus '|' which we want for tables).
        // See: https://spec.commonmark.org/ (backslash escapes). We keep the set small and pragmatic.
        return c switch {
            '\\' => true,
            '`' => true,
            '*' => true,
            '_' => true,
            '{' => true,
            '}' => true,
            '[' => true,
            ']' => true,
            '(' => true,
            ')' => true,
            '#' => true,
            '+' => true,
            '-' => true,
            '.' => true,
            '!' => true,
            '"' => true,
            '\'' => true,
            '|' => true,
            '>' => true,
            '=' => true,
            _ => false
        };
    }

    private static bool IsIntrawordDelimiter(string text, int start, int markerLength) {
        // Pragmatic GFM-ish rule: treat '_' emphasis markers as disabled when they appear inside "words".
        // This avoids accidentally italicizing identifiers like foo_bar_baz.
        if (string.IsNullOrEmpty(text)) return false;
        int left = start - 1;
        int right = start + markerLength;
        if (left < 0 || right >= text.Length) return false;
        return char.IsLetterOrDigit(text[left]) && char.IsLetterOrDigit(text[right]);
    }

    private static bool IsPotentialInlineStart(char c, bool allowInlineHtml, bool allowLinks, bool allowImages) {
        if (allowInlineHtml && c == '<') return true;
        if (c < PotentialInlineStartLookup.Length && PotentialInlineStartLookup[c]) {
            if (!allowLinks && c == '[') return false;
            if (!allowImages && c == '!') return false;
            return true;
        }
        return false;
    }

    private static bool TryParseSupportedInlineHtmlTag(
        string text,
        int start,
        MarkdownReaderOptions options,
        MarkdownReaderState? state,
        bool allowLinks,
        bool allowImages,
        out int consumed,
        out IMarkdownInline htmlNode) {
        consumed = 0;
        htmlNode = null!;

        if (string.IsNullOrEmpty(text) || start < 0 || start >= text.Length || text[start] != '<') {
            return false;
        }

        string[] tags = { "u", "sup", "sub", "ins", "q" };
        for (int i = 0; i < tags.Length; i++) {
            if (!TryParseInlineHtmlWrapper(text, start, tags[i], options, state, allowLinks, allowImages, out consumed, out var inlines)) {
                continue;
            }

            htmlNode = new HtmlTagSequenceInline(tags[i], inlines);
            return true;
        }

        return false;
    }

    private static bool TryParseInlineHtmlWrapper(
        string text,
        int start,
        string tagName,
        MarkdownReaderOptions options,
        MarkdownReaderState? state,
        bool allowLinks,
        bool allowImages,
        out int consumed,
        out InlineSequence inlines) {
        consumed = 0;
        inlines = new InlineSequence();

        if (!StartsWithExactHtmlTag(text, start, tagName, opening: true)) {
            return false;
        }

        int openLength = tagName.Length + 2;
        int scan = start + openLength;
        int depth = 1;

        while (scan < text.Length) {
            if (StartsWithExactHtmlTag(text, scan, tagName, opening: false)) {
                depth--;
                if (depth == 0) {
                    string inner = text.Substring(start + openLength, scan - (start + openLength));
                    inlines = ParseInlinesInternal(inner, options, state, allowLinks, allowImages);
                    DecodeHtmlEntitiesInTextRuns(inlines);
                    consumed = (scan - start) + tagName.Length + 3;
                    return true;
                }

                scan += tagName.Length + 3;
                continue;
            }

            if (StartsWithExactHtmlTag(text, scan, tagName, opening: true)) {
                depth++;
                scan += openLength;
                continue;
            }

            scan++;
        }

        return false;
    }

    private static bool DecodeHtmlEntitiesInTextRuns(InlineSequence sequence) {
        if (sequence == null || sequence.Nodes.Count == 0) {
            return false;
        }

        var rewritten = new List<IMarkdownInline>(sequence.Nodes.Count);
        bool changed = false;

        for (int i = 0; i < sequence.Nodes.Count; i++) {
            var node = sequence.Nodes[i];
            if (node == null) {
                continue;
            }

            rewritten.Add(DecodeHtmlEntitiesInInlineNode(node, ref changed));
        }

        if (changed) {
            sequence.ReplaceItems(rewritten);
        }

        return changed;
    }

    private static IMarkdownInline DecodeHtmlEntitiesInInlineNode(IMarkdownInline node, ref bool changed) {
        if (node is TextRun text) {
            string decoded = System.Net.WebUtility.HtmlDecode(text.Text);
            if (!string.Equals(decoded, text.Text, StringComparison.Ordinal)) {
                changed = true;
                return new DecodedHtmlEntityTextRun(decoded);
            }

            return text;
        }

        if (node is IInlineContainerMarkdownInline container && container.NestedInlines != null) {
            if (DecodeHtmlEntitiesInTextRuns(container.NestedInlines)) {
                changed = true;
            }
        }

        return node;
    }

    private static bool StartsWithExactHtmlTag(string text, int start, string tagName, bool opening) {
        if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(tagName) || start < 0 || start >= text.Length || text[start] != '<') {
            return false;
        }

        int position = start + 1;
        if (!opening) {
            if (position >= text.Length || text[position] != '/') {
                return false;
            }
            position++;
        }

        if (position + tagName.Length >= text.Length) {
            return false;
        }

        if (string.Compare(text, position, tagName, 0, tagName.Length, StringComparison.OrdinalIgnoreCase) != 0) {
            return false;
        }

        position += tagName.Length;
        if (position >= text.Length || text[position] != '>') {
            return false;
        }

        return true;
    }

    private static bool TryParseLink(
        string text,
        int start,
        out int consumed,
        out string label,
        out string href,
        out string? title,
        out int hrefStart,
        out int hrefLength,
        out int? titleStart,
        out int? titleLength) {
        consumed = 0; label = href = string.Empty; title = null;
        hrefStart = 0; hrefLength = 0; titleStart = null; titleLength = null;
        if (start >= text.Length || text[start] != '[') return false;
        int labelEnd = FindMatchingBracket(text, start);
        if (labelEnd < 0) return false;
        int parenOpen = (labelEnd + 1 < text.Length && text[labelEnd + 1] == '(') ? labelEnd + 1 : -1;
        if (parenOpen < 0) return false;
        int parenClose = FindMatchingParen(text, parenOpen);
        if (parenClose < 0) return false;
        label = text.Substring(start + 1, labelEnd - (start + 1));
        string inner = text.Substring(parenOpen + 1, parenClose - (parenOpen + 1));
        if (!TrySplitUrlAndOptionalTitle(inner, out href, out title, out int hrefInnerStart, out int hrefInnerLength, out int? titleInnerStart, out int? titleInnerLength)) {
            if (IndexOfWhitespace(inner.Trim()) >= 0) return false;
            href = UnescapeMarkdownBackslashEscapes(inner.Trim());
            title = null;
            int trimmedStart = 0;
            while (trimmedStart < inner.Length && char.IsWhiteSpace(inner[trimmedStart])) {
                trimmedStart++;
            }

            int trimmedEndExclusive = inner.Length;
            while (trimmedEndExclusive > trimmedStart && char.IsWhiteSpace(inner[trimmedEndExclusive - 1])) {
                trimmedEndExclusive--;
            }

            hrefInnerStart = trimmedStart;
            hrefInnerLength = Math.Max(0, trimmedEndExclusive - trimmedStart);
            titleInnerStart = null;
            titleInnerLength = null;
        }

        hrefStart = parenOpen + 1 + hrefInnerStart;
        hrefLength = hrefInnerLength;
        if (titleInnerStart.HasValue && titleInnerLength.HasValue) {
            titleStart = parenOpen + 1 + titleInnerStart.Value;
            titleLength = titleInnerLength.Value;
        }

        consumed = parenClose - start + 1;
        return true;
    }

    private static bool TryParseImageLink(string text, int start, out int consumed, out string alt, out string img, out string? imgTitle, out string href, out string? hrefTitle) =>
        TryParseImageLink(text, start, out consumed, out alt, out img, out imgTitle, out href, out hrefTitle, out _, out _, out _, out _, out _, out _, out _, out _, out _, out _);

    private static bool TryParseImageLink(
        string text,
        int start,
        out int consumed,
        out string alt,
        out string img,
        out string? imgTitle,
        out string href,
        out string? hrefTitle,
        out int altStart,
        out int altLength,
        out int imgStart,
        out int imgLength,
        out int? imgTitleStart,
        out int? imgTitleLength,
        out int hrefStart,
        out int hrefLength,
        out int? hrefTitleStart,
        out int? hrefTitleLength) {
        consumed = 0; alt = img = href = string.Empty; imgTitle = hrefTitle = null;
        altStart = altLength = imgStart = imgLength = hrefStart = hrefLength = 0;
        imgTitleStart = imgTitleLength = hrefTitleStart = hrefTitleLength = null;
        if (start >= text.Length || text[start] != '[') return false;
        if (start + 1 >= text.Length || text[start + 1] != '!') return false;
        if (start + 2 >= text.Length || text[start + 2] != '[') return false;
        int altEnd = FindMatchingBracket(text, start + 2);
        if (altEnd < 0) return false;
        if (altEnd + 1 >= text.Length || text[altEnd + 1] != '(') return false;
        int imgClose = FindMatchingParen(text, altEnd + 1);
        if (imgClose < 0) return false;
        altStart = start + 3;
        altLength = altEnd - altStart;
        alt = text.Substring(altStart, altLength);
        string inner = text.Substring(altEnd + 2, imgClose - (altEnd + 2));
        if (!TrySplitUrlAndOptionalTitle(inner, out img, out imgTitle, out int imgInnerStart, out int imgInnerLength, out int? imgTitleInnerStart, out int? imgTitleInnerLength)) {
            if (IndexOfWhitespace(inner.Trim()) >= 0) return false;
            img = UnescapeMarkdownBackslashEscapes(inner.Trim());
            imgTitle = null;
            int trimmedStart = 0;
            while (trimmedStart < inner.Length && char.IsWhiteSpace(inner[trimmedStart])) {
                trimmedStart++;
            }

            int trimmedEndExclusive = inner.Length;
            while (trimmedEndExclusive > trimmedStart && char.IsWhiteSpace(inner[trimmedEndExclusive - 1])) {
                trimmedEndExclusive--;
            }

            imgInnerStart = trimmedStart;
            imgInnerLength = Math.Max(0, trimmedEndExclusive - trimmedStart);
            imgTitleInnerStart = null;
            imgTitleInnerLength = null;
        }
        imgStart = altEnd + 2 + imgInnerStart;
        imgLength = imgInnerLength;
        imgTitleStart = imgTitleInnerStart.HasValue ? altEnd + 2 + imgTitleInnerStart.Value : null;
        imgTitleLength = imgTitleInnerLength;
        int closeBracket = (imgClose + 1 < text.Length && text[imgClose + 1] == ']') ? imgClose + 1 : -1;
        if (closeBracket < 0) return false;
        int parenOpen2 = (closeBracket + 1 < text.Length && text[closeBracket + 1] == '(') ? closeBracket + 1 : -1;
        if (parenOpen2 != closeBracket + 1) return false;
        int parenClose2 = FindMatchingParen(text, parenOpen2);
        if (parenClose2 < 0) return false;
        string hrefInner = text.Substring(parenOpen2 + 1, parenClose2 - (parenOpen2 + 1));
        if (!TrySplitUrlAndOptionalTitle(hrefInner, out href, out hrefTitle, out int hrefInnerStart, out int hrefInnerLength, out int? hrefTitleInnerStart, out int? hrefTitleInnerLength)) {
            if (IndexOfWhitespace(hrefInner.Trim()) >= 0) return false;
            href = UnescapeMarkdownBackslashEscapes(hrefInner.Trim());
            hrefTitle = null;
            int trimmedStart = 0;
            while (trimmedStart < hrefInner.Length && char.IsWhiteSpace(hrefInner[trimmedStart])) {
                trimmedStart++;
            }

            int trimmedEndExclusive = hrefInner.Length;
            while (trimmedEndExclusive > trimmedStart && char.IsWhiteSpace(hrefInner[trimmedEndExclusive - 1])) {
                trimmedEndExclusive--;
            }

            hrefInnerStart = trimmedStart;
            hrefInnerLength = Math.Max(0, trimmedEndExclusive - trimmedStart);
            hrefTitleInnerStart = null;
            hrefTitleInnerLength = null;
        }
        hrefStart = parenOpen2 + 1 + hrefInnerStart;
        hrefLength = hrefInnerLength;
        hrefTitleStart = hrefTitleInnerStart.HasValue ? parenOpen2 + 1 + hrefTitleInnerStart.Value : null;
        hrefTitleLength = hrefTitleInnerLength;
        consumed = parenClose2 - start + 1;
        return true;
    }

    private static bool TrySplitUrlAndOptionalTitle(
        string? inner,
        out string url,
        out string? title,
        out int urlStart,
        out int urlLength,
        out int? titleStart,
        out int? titleLength) {
        url = string.Empty;
        title = null;
        urlStart = 0;
        urlLength = 0;
        titleStart = null;
        titleLength = null;
        if (inner == null) return false;
        if (string.IsNullOrWhiteSpace(inner)) return false;

        int start = 0;
        while (start < inner.Length && char.IsWhiteSpace(inner[start])) {
            start++;
        }

        int endExclusive = inner.Length;
        while (endExclusive > start && char.IsWhiteSpace(inner[endExclusive - 1])) {
            endExclusive--;
        }

        if (endExclusive <= start) return false;

        // CommonMark: destination can be wrapped in <...> to allow spaces and parentheses safely.
        if (inner[start] == '<') {
            int gt = inner.IndexOf('>', start + 1);
            if (gt >= start + 1 && gt < endExclusive) {
                urlStart = start + 1;
                urlLength = gt - urlStart;
                url = UnescapeMarkdownBackslashEscapes(inner.Substring(urlStart, urlLength).Trim());

                int restStart = gt + 1;
                while (restStart < endExclusive && char.IsWhiteSpace(inner[restStart])) {
                    restStart++;
                }

                if (restStart >= endExclusive) {
                    return true;
                }

                if (!TryParseOptionalTitleToken(inner, restStart, endExclusive, out title, out int parsedTitleStart, out int parsedTitleLength)) {
                    return false;
                }

                title = UnescapeMarkdownBackslashEscapes(title!);
                titleStart = parsedTitleStart;
                titleLength = parsedTitleLength;
                return true;
            }
        }

        int ws = -1;
        for (int i = start; i < endExclusive; i++) {
            if (char.IsWhiteSpace(inner[i])) {
                ws = i;
                break;
            }
        }

        if (ws < 0) {
            urlStart = start;
            urlLength = endExclusive - start;
            url = UnescapeMarkdownBackslashEscapes(inner.Substring(urlStart, urlLength));
            title = null;
            return true;
        }

        urlStart = start;
        urlLength = ws - start;
        url = UnescapeMarkdownBackslashEscapes(inner.Substring(urlStart, urlLength).Trim());

        int remainingStart = ws;
        while (remainingStart < endExclusive && char.IsWhiteSpace(inner[remainingStart])) {
            remainingStart++;
        }

        if (remainingStart >= endExclusive) { title = null; return true; }

        if (!TryParseOptionalTitleToken(inner, remainingStart, endExclusive, out title, out int parsedStart, out int parsedLength)) return false;
        title = UnescapeMarkdownBackslashEscapes(title!);
        titleStart = parsedStart;
        titleLength = parsedLength;
        return true;
    }

    private static bool TrySplitUrlAndOptionalTitle(string? inner, out string url, out string? title) =>
        TrySplitUrlAndOptionalTitle(inner, out url, out title, out _, out _, out _, out _);

    private static int IndexOfWhitespace(string s) {
        for (int i = 0; i < s.Length; i++) if (char.IsWhiteSpace(s[i])) return i;
        return -1;
    }

    private static string? TryParseOptionalTitleToken(string s) {
        if (string.IsNullOrWhiteSpace(s)) return null;
        int start = 0;
        while (start < s.Length && char.IsWhiteSpace(s[start])) {
            start++;
        }

        int endExclusive = s.Length;
        while (endExclusive > start && char.IsWhiteSpace(s[endExclusive - 1])) {
            endExclusive--;
        }

        return TryParseOptionalTitleToken(s, start, endExclusive, out string? title, out _, out _) ? title : null;
    }

    private static bool TryParseOptionalTitleToken(
        string s,
        int start,
        int endExclusive,
        out string? title,
        out int titleStart,
        out int titleLength) {
        title = null;
        titleStart = 0;
        titleLength = 0;
        if (string.IsNullOrEmpty(s) || endExclusive - start < 2) {
            return false;
        }

        char open = s[start];
        char close = s[endExclusive - 1];
        if ((open == '"' && close == '"') ||
            (open == '\'' && close == '\'') ||
            (open == '(' && close == ')')) {
            titleStart = start + 1;
            titleLength = endExclusive - start - 2;
            title = s.Substring(titleStart, titleLength);
            return true;
        }

        return false;
    }

    private static bool TryParseInlineImage(string text, int start, out int consumed, out string alt, out string src, out string? title) =>
        TryParseInlineImage(
            text,
            start,
            out consumed,
            out alt,
            out src,
            out title,
            out _,
            out _,
            out _,
            out _,
            out _,
            out _);

    private static bool TryParseInlineImage(
        string text,
        int start,
        out int consumed,
        out string alt,
        out string src,
        out string? title,
        out int altStart,
        out int altLength,
        out int srcStart,
        out int srcLength,
        out int? titleStart,
        out int? titleLength) {
        consumed = 0;
        alt = src = string.Empty;
        title = null;
        altStart = altLength = srcStart = srcLength = 0;
        titleStart = titleLength = null;
        if (start + 1 >= text.Length || text[start] != '!' || text[start + 1] != '[') return false;
        int altEnd = FindMatchingBracket(text, start + 1);
        if (altEnd < 0) return false;
        if (altEnd + 1 >= text.Length || text[altEnd + 1] != '(') return false;
        int parenClose = FindMatchingParen(text, altEnd + 1);
        if (parenClose < 0) return false;
        altStart = start + 2;
        altLength = altEnd - altStart;
        alt = text.Substring(altStart, altLength);
        string inner = text.Substring(altEnd + 2, parenClose - (altEnd + 2));
        if (!TrySplitUrlAndOptionalTitle(
            inner,
            out src,
            out title,
            out int srcInnerStart,
            out int srcInnerLength,
            out int? titleInnerStart,
            out int? titleInnerLength)) {
            if (IndexOfWhitespace(inner.Trim()) >= 0) return false;
            src = UnescapeMarkdownBackslashEscapes(inner.Trim());
            title = null;

            int trimmedStart = 0;
            while (trimmedStart < inner.Length && char.IsWhiteSpace(inner[trimmedStart])) {
                trimmedStart++;
            }

            int trimmedEndExclusive = inner.Length;
            while (trimmedEndExclusive > trimmedStart && char.IsWhiteSpace(inner[trimmedEndExclusive - 1])) {
                trimmedEndExclusive--;
            }

            srcInnerStart = trimmedStart;
            srcInnerLength = Math.Max(0, trimmedEndExclusive - trimmedStart);
            titleInnerStart = null;
            titleInnerLength = null;
        }

        srcStart = altEnd + 2 + srcInnerStart;
        srcLength = srcInnerLength;
        titleStart = titleInnerStart.HasValue ? altEnd + 2 + titleInnerStart.Value : null;
        titleLength = titleInnerLength;
        consumed = parenClose - start + 1;
        return true;
    }

    private static bool TryParseReferenceImage(string text, int start, out int consumed, out string alt, out string label) =>
        TryParseReferenceImage(text, start, out consumed, out alt, out label, out _, out _);

    private static bool TryParseReferenceImage(
        string text,
        int start,
        out int consumed,
        out string alt,
        out string label,
        out int altStart,
        out int altLength) {
        consumed = 0;
        alt = label = string.Empty;
        altStart = altLength = 0;
        if (start + 1 >= text.Length || text[start] != '!' || text[start + 1] != '[') return false;
        int altEnd = FindMatchingBracket(text, start + 1);
        if (altEnd < 0) return false;

        altStart = start + 2;
        altLength = altEnd - altStart;
        alt = text.Substring(altStart, altLength);

        // Inline image uses "(...)" and is handled elsewhere.
        if (altEnd + 1 < text.Length && text[altEnd + 1] == '(') return false;

        // Full or collapsed reference: ![alt][label] or ![alt][]
        if (altEnd + 1 < text.Length && text[altEnd + 1] == '[') {
            int labelEnd = FindMatchingBracket(text, altEnd + 1);
            if (labelEnd < 0) return false;
            label = text.Substring(altEnd + 2, labelEnd - (altEnd + 2));
            if (string.IsNullOrEmpty(label)) label = alt;
            consumed = labelEnd - start + 1;
            return true;
        }

        // Shortcut: ![label]
        label = alt;
        consumed = altEnd - start + 1;
        return true;
    }

    private static int FindMatchingParen(string text, int openIndex) {
        int depth = 0;
        bool inDoubleQuotes = false;
        bool inSingleQuotes = false;
        bool inAngle = false;
        bool escaped = false;
        for (int i = openIndex; i < text.Length; i++) {
            char c = text[i];
            if (escaped) {
                escaped = false;
                continue;
            }
            if (c == '\\') {
                escaped = true;
                continue;
            }
            if (inAngle) {
                if (c == '>') inAngle = false;
                continue;
            }
            if (inDoubleQuotes) {
                if (c == '"') inDoubleQuotes = false;
                continue;
            }
            if (inSingleQuotes) {
                if (c == '\'') inSingleQuotes = false;
                continue;
            }
            if (c == '(') { depth++; continue; }
            if (c == ')') { depth--; if (depth == 0) return i; continue; }
            if (depth == 1) {
                if (c == '<') { inAngle = true; continue; }
                if (c == '"') { inDoubleQuotes = true; continue; }
                if (c == '\'') { inSingleQuotes = true; continue; }
            }
        }
        return -1;
    }

    private static bool StartsWithHttp(string text, int start, out int end) {
        end = start;
        if (start + 7 > text.Length) return false;
        // Require a boundary on the left so we don't linkify inside longer words.
        if (HasInvalidAutolinkLeftBoundary(text, start)) return false;
        if (IsAfterInvalidReferenceDefinitionPrefix(text, start)) return false;
        var rem = text.Substring(start);
        if (!(rem.StartsWith("http://") || rem.StartsWith("https://"))) return false;
        int rawEnd = ConsumeLiteralUrl(text, start);
        int i = rawEnd;
        // Trim trailing punctuation commonly outside URLs
        while (i > start && (text[i - 1] == '.' || text[i - 1] == ',' || text[i - 1] == ';' || text[i - 1] == ':' || text[i - 1] == '!' || text[i - 1] == '?' || text[i - 1] == '\'' || text[i - 1] == '"')) i--;
        if (ShouldRejectQueryFragmentSpecialCharsAutolink(text, start, i)) return false;
        if (ShouldRejectAmbiguousTrailingParen(text, start, rawEnd, i)) return false;
        end = i; return end > start + 7;
    }

    private static bool StartsWithWww(string text, int start, out int end) {
        end = start;
        if (start + 4 > text.Length) return false;
        if (HasInvalidAutolinkLeftBoundary(text, start)) return false;
        if (IsAfterInvalidReferenceDefinitionPrefix(text, start)) return false;
        if (!(text.Substring(start).StartsWith("www.", StringComparison.OrdinalIgnoreCase))) return false;

        int rawEnd = ConsumeLiteralUrl(text, start);
        int i = rawEnd;
        int scanEnd = rawEnd;
        while (i > start && (text[i - 1] == '.' || text[i - 1] == ',' || text[i - 1] == ';' || text[i - 1] == ':' || text[i - 1] == '!' || text[i - 1] == '?' || text[i - 1] == '\'' || text[i - 1] == '"')) i--;
        if (ShouldRejectQueryFragmentSpecialCharsAutolink(text, start, i)) return false;
        if (ShouldRejectAmbiguousTrailingParen(text, start, rawEnd, i)) return false;

        // Must include at least one dot after the www.
        var token = text.Substring(start, i - start);
        if (token.Length <= 4) return false;
        if (token.IndexOf('.', 4) < 0) return false;

        // Right boundary: avoid linking as part of an identifier-like token.
        if (scanEnd < text.Length && IsEmailChar(text[scanEnd])) return false;

        end = i;
        return end > start + 4;
    }

    private static bool HasInvalidAutolinkLeftBoundary(string text, int start) {
        if (string.IsNullOrEmpty(text) || start <= 0 || start > text.Length) return false;

        char previous = text[start - 1];
        return char.IsLetterOrDigit(previous)
            || previous == '_'
            || previous == '/'
            || previous == ':'
            || previous == '.'
            || previous == '+'
            || previous == '-'
            || previous == '='
            || previous == '&'
            || previous == '('
            || previous == '\''
            || previous == '[';
    }

    private static bool IsAfterInvalidReferenceDefinitionPrefix(string text, int start) {
        if (string.IsNullOrEmpty(text) || start <= 0 || start > text.Length) return false;

        int lineStart = text.LastIndexOf('\n', start - 1);
        lineStart = lineStart < 0 ? 0 : lineStart + 1;
        int lineEnd = text.IndexOf('\n', start);
        if (lineEnd < 0) lineEnd = text.Length;

        string line = text.Substring(lineStart, lineEnd - lineStart);
        if (!StartsWithReferenceDefinitionLikeLabel(line)) return false;

        return !TryParseReferenceLinkDefinition(new[] { line }, 0, new MarkdownReaderOptions(), out _, out _, out _, out _);
    }

    private static int ConsumeLiteralUrl(string text, int start) {
        int i = start;
        int parenDepth = 0;
        while (i < text.Length) {
            char c = text[i];
            if (char.IsWhiteSpace(c)) break;
            if (c == ']' || c == '<') break;
            if (c == '(') {
                parenDepth++;
                i++;
                continue;
            }
            if (c == ')') {
                if (parenDepth == 0) break;
                parenDepth--;
                i++;
                continue;
            }
            i++;
        }

        return i;
    }

    private static bool ShouldRejectAmbiguousTrailingParen(string text, int start, int rawEnd, int trimmedEnd) {
        if (string.IsNullOrEmpty(text) || start < 0 || trimmedEnd <= start) return false;

        bool extraClosingParenOutsideUrl = rawEnd < text.Length && text[rawEnd] == ')';
        bool trailingPunctuationTrimmedAfterBalancedParen = rawEnd > trimmedEnd && text[trimmedEnd - 1] == ')';
        if (!extraClosingParenOutsideUrl && !trailingPunctuationTrimmedAfterBalancedParen) return false;
        if (start > 0 && text[start - 1] == '(') return false;

        bool sawOpenParen = false;
        for (int i = start; i < trimmedEnd - 1; i++) {
            if (text[i] == '(') {
                sawOpenParen = true;
                break;
            }
        }

        return sawOpenParen;
    }

    private static bool ShouldRejectQueryFragmentSpecialCharsAutolink(string text, int start, int end) {
        if (string.IsNullOrEmpty(text) || start < 0 || end <= start) return false;

        int queryOrFragmentIndex = -1;
        for (int i = start; i < end; i++) {
            char ch = text[i];
            if (ch == '?' || ch == '#') {
                queryOrFragmentIndex = i;
                break;
            }
        }

        if (queryOrFragmentIndex < 0) return false;

        for (int i = queryOrFragmentIndex + 1; i < end; i++) {
            char ch = text[i];
            if (ch == '(' || ch == ')' || ch == '&') {
                return true;
            }
        }

        return false;
    }

    private static bool TryConsumePlainEmail(string text, int start, out int end, out string email) {
        end = start;
        email = string.Empty;
        if (start < 0 || start >= text.Length) return false;
        if (!IsEmailStartChar(text[start])) return false;
        if (start > 0 && (IsEmailChar(text[start - 1]) || text[start - 1] == '+' || text[start - 1] == '/' || text[start - 1] == ':' || text[start - 1] == '=' || text[start - 1] == '&' || text[start - 1] == '(' || text[start - 1] == '\'' || text[start - 1] == '[')) return false;
        if (IsImmediatelyAfterMailtoScheme(text, start)) return false;

        int i = start;
        bool sawAt = false;
        // Stop at whitespace or common "outside token" delimiters; keep it pragmatic.
        while (i < text.Length) {
            char c = text[i];
            if (char.IsWhiteSpace(c)) break;
            if (c == ')' || c == ']' || c == '<') break;
            if (!IsEmailChar(c)) break;
            if (c == '@') sawAt = true;
            i++;
        }
        if (!sawAt) return false;

        int scanEnd = i;
        int j = i;
        while (j > start && (text[j - 1] == '.' || text[j - 1] == ',' || text[j - 1] == ';' || text[j - 1] == ':')) j--;
        if (j <= start) return false;

        var token = text.Substring(start, j - start);
        if (!LooksLikeEmail(token)) return false;
        if (scanEnd < text.Length) {
            if (IsEmailChar(text[scanEnd])) return false;
            if (text[scanEnd] == '/' || text[scanEnd] == '#') return false;
        }

        end = j;
        email = token;
        return true;
    }

    private static bool IsEmailStartChar(char c) => char.IsLetterOrDigit(c);

    private static bool IsEmailChar(char c) {
        if (char.IsLetterOrDigit(c)) return true;
        return c == '@' || c == '.' || c == '-' || c == '_';
    }

    private static bool IsImmediatelyAfterMailtoScheme(string text, int start) {
        if (string.IsNullOrEmpty(text) || start < 7) return false;
        if (text[start - 1] != ':') return false;

        return string.Compare(text, start - 7, "mailto:", 0, 7, StringComparison.OrdinalIgnoreCase) == 0;
    }

    /// <summary>
    /// Parses a single line of Markdown inline content into a typed <see cref="InlineSequence"/>.
    /// This helper is exposed to allow other components (e.g., Word converter) to interpret
    /// inline markup in contexts like table cells where we currently store raw strings.
    /// </summary>
    /// <param name="text">Inline Markdown text.</param>
    /// <param name="options">Reader options controlling inline interpretation.</param>
    /// <returns>Parsed sequence of inline nodes.</returns>
    public static InlineSequence ParseInlineText(string? text, MarkdownReaderOptions? options = null) =>
        ParseInlineText(text, options, null);

    internal static InlineSequence ParseInlineText(string? text, MarkdownReaderOptions? options, MarkdownReaderState? state) =>
        ParseInlines(text ?? string.Empty, options ?? new MarkdownReaderOptions(), state);

    private static IReadOnlyList<MarkdownInlineParserExtension> BuildEffectiveInlineParserExtensions(MarkdownReaderOptions options) {
        if (options.InlineParserExtensions.Count == 0) {
            return Array.Empty<MarkdownInlineParserExtension>();
        }

        var active = new List<MarkdownInlineParserExtension>(options.InlineParserExtensions.Count);
        for (var i = 0; i < options.InlineParserExtensions.Count; i++) {
            var extension = options.InlineParserExtensions[i];
            if (extension != null && extension.AppliesTo(options)) {
                active.Add(extension);
            }
        }

        return active;
    }

    private static bool TryParseInlineExtension(
        string text,
        int position,
        MarkdownReaderOptions options,
        MarkdownReaderState? state,
        bool allowLinks,
        bool allowImages,
        MarkdownInlineSourceMap? sourceMap,
        IReadOnlyList<MarkdownInlineParserExtension> inlineParserExtensions,
        Func<int, int, bool, bool, InlineSequence> parseNestedInlineSegment,
        out MarkdownInlineParseResult result) {
        result = default;
        if (inlineParserExtensions.Count == 0) {
            return false;
        }

        var context = new MarkdownInlineParserContext(
            text,
            position,
            options,
            state,
            allowLinks,
            allowImages,
            sourceMap,
            parseNestedInlineSegment);

        for (var i = 0; i < inlineParserExtensions.Count; i++) {
            var extension = inlineParserExtensions[i];
            if (!extension.Parser(context, out result)) {
                continue;
            }

            if (result.ConsumedLength <= 0) {
                throw new InvalidOperationException($"Inline parser extension '{extension.Name}' returned a non-positive consumed length.");
            }

            if (position + result.ConsumedLength > text.Length) {
                throw new InvalidOperationException($"Inline parser extension '{extension.Name}' consumed past the end of the input.");
            }

            return true;
        }

        result = default;
        return false;
    }
}
