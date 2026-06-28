namespace OfficeIMO.Markdown;

/// <summary>
/// Inline parsing helpers for <see cref="MarkdownReader"/>.
/// </summary>
public static partial class MarkdownReader {
    private static InlineSequence ParseInlines(string text, MarkdownReaderOptions options, MarkdownReaderState? state = null, MarkdownInlineSourceMap? sourceMap = null) {
        var sequence = ParseInlinesInternal(text, options, state, allowLinks: true, allowImages: true, sourceMap);
        NormalizeInlineSequenceInPlace(sequence, options.InputNormalization);
        ApplyInlineTransformExtensions(sequence, text, options);
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
        void AddCodeSpanNode(string literal, int start, int closingStart, int fenceLength) {
            var node = new CodeSpanInline(literal);
            var marker = new string('`', fenceLength);
            var contentStart = start + fenceLength;
            var contentLength = closingStart - contentStart;
            if (contentLength > 0) {
                MarkdownInlineMetadataSourceSpans.SetCodeSpanContent(
                    node,
                    sourceMap?.GetSpan(contentStart, contentLength));
            }

            MarkdownInlineMetadataSourceSpans.SetFormattingMarkers(
                node,
                marker,
                sourceMap?.GetSpan(start, fenceLength),
                marker,
                sourceMap?.GetSpan(closingStart, fenceLength));
            AddRawNode(node, start, closingStart + fenceLength - start);
        }
        void AddTextNode(string literal, int start, int length) => AddRawNode(new TextRun(literal), start, length);
        void AddEscapedTextNode(char escapedCharacter, int start) {
            var node = new TextRun(escapedCharacter.ToString());
            MarkdownInlineMetadataSourceSpans.SetEscapedText(
                node,
                "\\",
                sourceMap?.GetSpan(start, 1),
                escapedCharacter.ToString(),
                sourceMap?.GetSpan(start + 1, 1));
            AddRawNode(node, start, 2);
        }
        void AddDecodedHtmlEntityNode(string literal, int start, int length) {
            var node = new DecodedHtmlEntityTextRun(literal);
            MarkdownInlineMetadataSourceSpans.SetDecodedEntity(
                node,
                text.Substring(start, length),
                sourceMap?.GetSpan(start, length));
            AddRawNode(node, start, length);
        }
        void AddHardBreakNode(int start, int length) {
            var node = new HardBreakInline();
            var tokenLiteral = sourceMap?.GetTokenLiteral(start, length);
            var isSyntheticSoftBreak = length == 1 && start >= 0 && start < text.Length && text[start] == '\n' && tokenLiteral == null;
            if (!isSyntheticSoftBreak) {
                MarkdownInlineMetadataSourceSpans.SetHardBreakMarker(
                    node,
                    tokenLiteral ?? text.Substring(start, length),
                    sourceMap?.GetSpan(start, length));
                AddRawNode(node, start, length);
                return;
            }

            Current().AddRaw(node);
        }
        void AddFootnoteRefNode(string label, int start, int length) {
            var node = new FootnoteRefInline(label);
            MarkdownInlineMetadataSourceSpans.SetFormattingMarkers(
                node,
                "[^",
                sourceMap?.GetSpan(start, 2),
                "]",
                sourceMap?.GetSpan(start + length - 1, 1));
            AddRawNode(node, start, length);
        }
        void AddAbbreviationNode(MarkdownAbbreviationDefinition definition, int start) {
            var node = new AbbreviationInline(definition.Label, definition.Title);
            MarkdownInlineMetadataSourceSpans.SetAbbreviationParts(
                node,
                sourceMap?.GetSpan(start, definition.Label.Length),
                definition.TitleSourceSpan);
            AddRawNode(node, start, definition.Label.Length);
        }
        void AddAutolinkNode(
            string label,
            string resolvedHref,
            int start,
            int length,
            int targetStart,
            int targetLength,
            bool angleWrapped) {
            var link = new LinkInline(label, resolvedHref, null);
            AddRawNode(link, start, length);
            MarkdownInlineMetadataSourceSpans.SetLinkParts(
                link,
                sourceMap?.GetSpan(targetStart, targetLength),
                null);

            if (angleWrapped) {
                MarkdownInlineMetadataSourceSpans.SetFormattingMarkers(
                    link,
                    "<",
                    sourceMap?.GetSpan(start, 1),
                    ">",
                    sourceMap?.GetSpan(start + length - 1, 1));
            }
        }
        void AddInlineLinkNode(
            InlineSequence label,
            string resolvedHref,
            string? title,
            int start,
            int length,
            int labelLength,
            int targetStart,
            int targetLength,
            int? titleStart,
            int? titleLength) {
            var link = new LinkInline(label, resolvedHref, title);
            AddRawNode(link, start, length);
            MarkdownInlineMetadataSourceSpans.SetLinkParts(
                link,
                targetLength > 0 ? sourceMap?.GetSpan(targetStart, targetLength) : null,
                titleStart.HasValue && titleLength.HasValue ? sourceMap?.GetSpan(titleStart.Value, titleLength.Value) : null);
            MarkdownInlineMetadataSourceSpans.SetFormattingMarkers(
                link,
                "[",
                sourceMap?.GetSpan(start, 1),
                ")",
                sourceMap?.GetSpan(start + length - 1, 1),
                "](",
                sourceMap?.GetSpan(start + labelLength + 1, 2));
        }
        void AddInlineImageNode(
            string alt,
            string resolvedSource,
            string? title,
            string plainAlt,
            int start,
            int length,
            int altStart,
            int altLength,
            int sourceStart,
            int sourceLength,
            int? titleStart,
            int? titleLength) {
            var image = new ImageInline(alt, resolvedSource, title, plainAlt);
            AddRawNode(image, start, length);
            MarkdownInlineMetadataSourceSpans.SetImageParts(
                image,
                sourceMap?.GetSpan(altStart, altLength),
                sourceMap?.GetSpan(sourceStart, sourceLength),
                titleStart.HasValue && titleLength.HasValue
                    ? sourceMap?.GetSpan(titleStart.Value, titleLength.Value)
                    : null);
            MarkdownInlineMetadataSourceSpans.SetFormattingMarkers(
                image,
                "![",
                sourceMap?.GetSpan(start, 2),
                ")",
                sourceMap?.GetSpan(start + length - 1, 1),
                "](",
                sourceMap?.GetSpan(altStart + altLength, 2));
        }
        void AddReferenceImageNode(
            string alt,
            string resolvedSource,
            string? title,
            string plainAlt,
            int start,
            int length,
            int altStart,
            int altLength,
            MarkdownSourceSpan? sourceSpan,
            MarkdownSourceSpan? titleSpan) {
            var image = new ImageInline(alt, resolvedSource, title, plainAlt);
            AddRawNode(image, start, length);
            MarkdownInlineMetadataSourceSpans.SetImageParts(
                image,
                sourceMap?.GetSpan(altStart, altLength),
                sourceSpan,
                titleSpan);
            int separatorStart = altStart + altLength;
            bool hasSeparator = HasReferenceSeparatorMarker(separatorStart, start, length);
            MarkdownInlineMetadataSourceSpans.SetFormattingMarkers(
                image,
                "![",
                sourceMap?.GetSpan(start, 2),
                "]",
                sourceMap?.GetSpan(start + length - 1, 1),
                hasSeparator ? "][" : null,
                hasSeparator ? sourceMap?.GetSpan(separatorStart, 2) : null);
        }
        void AddReferenceLinkNode(
            InlineSequence label,
            string resolvedHref,
            string? title,
            int start,
            int length,
            int labelLength,
            MarkdownSourceSpan? targetSpan,
            MarkdownSourceSpan? titleSpan) {
            var link = new LinkInline(label, resolvedHref, title);
            AddRawNode(link, start, length);
            MarkdownInlineMetadataSourceSpans.SetLinkParts(link, targetSpan, titleSpan);
            int separatorStart = start + labelLength + 1;
            bool hasSeparator = HasReferenceSeparatorMarker(separatorStart, start, length);
            MarkdownInlineMetadataSourceSpans.SetFormattingMarkers(
                link,
                "[",
                sourceMap?.GetSpan(start, 1),
                "]",
                sourceMap?.GetSpan(start + length - 1, 1),
                hasSeparator ? "][" : null,
                hasSeparator ? sourceMap?.GetSpan(separatorStart, 2) : null);
        }
        bool HasReferenceSeparatorMarker(int separatorStart, int start, int length) =>
            separatorStart + 1 < start + length
            && separatorStart + 1 < text.Length
            && text[separatorStart] == ']'
            && text[separatorStart + 1] == '[';
        int FindLinkedImageOuterSeparatorStart(int start, int length, int targetStart) {
            for (int i = Math.Min(targetStart - 1, start + length - 1); i > start; i--) {
                if (char.IsWhiteSpace(text[i])) {
                    continue;
                }

                return text[i] == '(' && i > start && text[i - 1] == ']' ? i - 1 : -1;
            }

            return -1;
        }
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

            if (options.Abbreviations
                && allowLinks
                && TryConsumeAbbreviation(text, pos, state, out var abbreviation)) {
                AddAbbreviationNode(abbreviation, pos);
                pos += abbreviation.Label.Length;
                continue;
            }

            // Backslash escape (CommonMark-ish): only escape punctuation we care about so that Windows paths like
            // "C:\Support\GitHub" keep their backslashes.
            if (text[pos] == '\\') {
                if (pos + 1 < text.Length) {
                    char next = text[pos + 1];
                    if (IsBackslashEscapable(next)) {
                        AddEscapedTextNode(next, pos);
                        pos += 2;
                        continue;
                    }
                }
                AddTextNode("\\", pos, 1);
                pos++;
                continue;
            }

            if (TryConsumeHtmlEntityText(text, pos, out int consumedEntity, out string decodedEntity)) {
                AddDecodedHtmlEntityNode(decodedEntity, pos, consumedEntity);
                pos += consumedEntity;
                continue;
            }

            // Autolink: http(s)://... until whitespace or closing punct
            if (options.AutolinkUrls && StartsWithHttp(text, pos, options, out int urlEnd)) {
                var url = text.Substring(pos, urlEnd - pos);
                var resolved = ResolveUrl(url, options);
                if (resolved is null) {
                    AddTextNode(url, pos, urlEnd - pos);
                } else {
                    AddAutolinkNode(url, resolved!, pos, urlEnd - pos, pos, urlEnd - pos, angleWrapped: false);
                }
                pos = urlEnd; continue;
            }

            // Autolink: www.example.com
            if (options.AutolinkWwwUrls && StartsWithWww(text, pos, options, out int wwwEnd)) {
                var label = text.Substring(pos, wwwEnd - pos);
                var scheme = string.IsNullOrWhiteSpace(options.AutolinkWwwScheme) ? "https://" : options.AutolinkWwwScheme.Trim();
                if (!scheme.EndsWith("://", StringComparison.Ordinal)) scheme = scheme.TrimEnd('/') + "://";
                var href = scheme + label;
                var resolved = ResolveUrl(href, options);
                if (resolved is null) {
                    AddTextNode(label, pos, wwwEnd - pos);
                } else {
                    AddAutolinkNode(label, resolved!, pos, wwwEnd - pos, pos, wwwEnd - pos, angleWrapped: false);
                }
                pos = wwwEnd; continue;
            }

            // Autolink: GFM URI-like bare schemes such as mailto: and xmpp:
            if (options.AutolinkBareSchemeUrls && TryConsumeBareSchemeAutolink(text, pos, options, out int schemeEnd, out string schemeLabel, out string schemeHref)) {
                var resolved = ResolveUrl(schemeHref, options);
                if (resolved is null) {
                    AddTextNode(schemeLabel, pos, schemeEnd - pos);
                } else {
                    AddAutolinkNode(schemeLabel, resolved!, pos, schemeEnd - pos, pos, schemeEnd - pos, angleWrapped: false);
                }
                pos = schemeEnd; continue;
            }

            // Autolink: plain email
            if (options.AutolinkEmails && TryConsumePlainEmail(text, pos, options, out int emailEnd, out string email)) {
                var href = "mailto:" + email;
                var resolved = ResolveUrl(href, options);
                if (resolved is null) {
                    AddTextNode(email, pos, emailEnd - pos);
                } else {
                    AddAutolinkNode(email, resolved!, pos, emailEnd - pos, pos, emailEnd - pos, angleWrapped: false);
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
                    AddCodeSpanNode(inner, pos, matchStart, fenceLen);
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
                if (rb > pos + 2) { var lab = text.Substring(pos + 2, rb - (pos + 2)); AddFootnoteRefNode(lab, pos, rb + 1 - pos); pos = rb + 1; continue; }
            }

            if (TryParseImageLink(
                text,
                pos,
                sourceMap,
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
                        int outerSeparatorStart = FindLinkedImageOuterSeparatorStart(pos, consumed, imageLinkHrefStart);
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
                        MarkdownInlineMetadataSourceSpans.SetFormattingMarkers(
                            imageLink,
                            "[",
                            sourceMap?.GetSpan(pos, 1),
                            ")",
                            sourceMap?.GetSpan(pos + consumed - 1, 1),
                            outerSeparatorStart >= 0 ? "](" : null,
                            outerSeparatorStart >= 0 ? sourceMap?.GetSpan(outerSeparatorStart, 2) : null);
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
                                AddReferenceImageNode(
                                    altRef,
                                    resolved!,
                                    defImg.Title,
                                    plainAltRef,
                                    pos,
                                    consumedRefImg,
                                    altRefStart,
                                    altRefLength,
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
                        sourceMap,
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
                            AddInlineImageNode(
                                altImg,
                                srcResolved!,
                                titleImg,
                                plainAltImg,
                                pos,
                                consumedImg,
                                altStartImg,
                                altLengthImg,
                                srcStartImg,
                                srcLengthImg,
                                titleStartImg,
                                titleLengthImg);
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
                    AddAutolinkNode(
                        labelAngle,
                        resolved!,
                        pos,
                        consumedAngle,
                        pos + 1,
                        Math.Max(0, consumedAngle - 2),
                        angleWrapped: true);
                }
                pos += consumedAngle;
                continue;
            }
            if (text[pos] == '[') {
                if (allowLinks) {
                    if (TryParseLink(text, pos, sourceMap, state, out int consumed2, out var label2, out var href3, out var title2, out int hrefStart2, out int hrefLength2, out int? titleStart2, out int? titleLength2)) {
                        var labelSeq = ParseInlinesInternal(label2, options, state, allowLinks: false, allowImages: true, SliceMap(pos + 1, label2.Length));

                        // Allow empty href: commonly used as placeholder or to be filled by the host.
                        if (string.IsNullOrWhiteSpace(href3)) {
                            AddInlineLinkNode(labelSeq, string.Empty, title2, pos, consumed2, label2.Length, hrefStart2, hrefLength2, titleStart2, titleLength2);
                        } else {
                            var hrefResolved = ResolveUrl(href3, options);
                            if (hrefResolved is null) {
                                // Unsafe URLs: keep the label as plain inline content instead of producing an <a href="...">.
                                foreach (var n in labelSeq.Nodes) Current().AddRaw(n);
                            } else {
                                AddInlineLinkNode(labelSeq, hrefResolved!, title2, pos, consumed2, label2.Length, hrefStart2, hrefLength2, titleStart2, titleLength2);
                            }
                        }
                        pos += consumed2; continue;
                    }

                    if (state != null && TryParseCollapsedRef(text, pos, out int consumedC, out var lbl2)) {
                        var key = NormalizeReferenceLabel(lbl2);
                        if (ContainsResolvedLinkInLabel(lbl2, state)) {
                            AddTextNode("[", pos, 1);
                            pos++;
                            continue;
                        }

                        var labelSeq = ParseInlinesInternal(lbl2, options, state, allowLinks: false, allowImages: true, SliceMap(pos + 1, lbl2.Length));
                        if (state.LinkRefs.TryGetValue(key, out var def2)) {
                            var resolved = ResolveUrl(def2.Url, options);
                            if (resolved is null) {
                                foreach (var n in labelSeq.Nodes) Current().AddRaw(n);
                            } else {
                                AddReferenceLinkNode(labelSeq, resolved!, def2.Title, pos, consumedC, lbl2.Length, def2.UrlSourceSpan, def2.TitleSourceSpan);
                            }
                            pos += consumedC; continue;
                        }

                        AddTextNode(text.Substring(pos, consumedC), pos, consumedC);
                        pos += consumedC; continue;
                    }
                    if (state != null && TryParseRefLink(text, pos, out int consumedR, out var lbl, out var refLabel)) {
                        var key = NormalizeReferenceLabel(refLabel);
                        if (ContainsResolvedLinkInLabel(lbl, state)) {
                            AddTextNode("[", pos, 1);
                            pos++;
                            continue;
                        }

                        var labelSeq = ParseInlinesInternal(lbl, options, state, allowLinks: false, allowImages: true, SliceMap(pos + 1, lbl.Length));
                        if (state.LinkRefs.TryGetValue(key, out var def)) {
                            var resolved = ResolveUrl(def.Url, options);
                            if (resolved is null) {
                                foreach (var n in labelSeq.Nodes) Current().AddRaw(n);
                            } else {
                                AddReferenceLinkNode(labelSeq, resolved!, def.Title, pos, consumedR, lbl.Length, def.UrlSourceSpan, def.TitleSourceSpan);
                            }
                            pos += consumedR; continue;
                        }
                    }
                    if (state != null && TryParseShortcutRef(text, pos, out int consumedS, out var lbl3)) {
                        var key = NormalizeReferenceLabel(lbl3);
                        if (ContainsResolvedLinkInLabel(lbl3, state)) {
                            AddTextNode("[", pos, 1);
                            pos++;
                            continue;
                        }

                        var labelSeq = ParseInlinesInternal(lbl3, options, state, allowLinks: false, allowImages: true, SliceMap(pos + 1, lbl3.Length));
                        if (state.LinkRefs.TryGetValue(key, out var def3)) {
                            var resolved = ResolveUrl(def3.Url, options);
                            if (resolved is null) {
                                foreach (var n in labelSeq.Nodes) Current().AddRaw(n);
                            } else {
                                AddReferenceLinkNode(labelSeq, resolved!, def3.Title, pos, consumedS, lbl3.Length, def3.UrlSourceSpan, def3.TitleSourceSpan);
                            }
                            pos += consumedS; continue;
                        }

                        if (stack.Count > 1 || ContainsBackslashEscapableCharacter(lbl3)) {
                            AddTextNode("[", pos, 1);
                            pos++;
                            continue;
                        }

                        AddTextNode(text.Substring(pos, consumedS), pos, consumedS);
                        pos += consumedS;
                        continue;
                    }
                }
            }

            // Emphasis / strong / strike / highlight / inserted / superscript / subscript using delimiter-run rules + an open-frame stack.
            if (text[pos] == '*' || text[pos] == '_' || text[pos] == '~' || text[pos] == '=' || text[pos] == '+' || text[pos] == '^') {
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

                // "==" and "++" always require a double delimiter. "~" can mean Markdig subscript or cmark-gfm single-tilde strike.
                if ((marker == '=' || marker == '+') && runLen < 2) {
                    AddTextNode(marker.ToString(), pos, 1);
                    pos++;
                    continue;
                }

                if (marker == '+' && runLen > 2) {
                    AddTextNode(new string(marker, runLen), pos, runLen);
                    pos += runLen;
                    continue;
                }

                if (marker == '~' && runLen < (options.SingleTildeStrikethrough || options.Subscript ? 1 : 2)) {
                    AddTextNode(marker.ToString(), pos, 1);
                    pos++;
                    continue;
                }

                // cmark-gfm strikethrough recognizes one- and two-tilde delimiter runs.
                // Longer runs such as "~~~three~~~" and "~~~~~one~~~~~" remain literal.
                if (marker == '~' && runLen > 2) {
                    AddTextNode(new string(marker, runLen), pos, runLen);
                    pos += runLen;
                    continue;
                }

                GetDelimiterFlags(text, pos, marker, runLen, out bool canOpen, out bool canClose);

                if (ShouldTreatMixedSingleMarkerAsLiteral(text, pos, marker, runLen, canOpen, canClose, stack)) {
                    AddTextNode(marker.ToString(), pos, 1);
                    pos++;
                    continue;
                }

                if (ShouldTreatOppositeMarkerBeforeOuterCloseAsLiteral(text, pos, marker, runLen, stack)) {
                    AddTextNode(new string(marker, runLen), pos, runLen);
                    pos += runLen;
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
                bool splitDoubleRunIntoRootDualItalic = ShouldSplitDoubleRunIntoRootDualItalic(text, pos, marker, remaining, canOpen, canClose, stack);
                int literalPrefixForOddCloser = GetLiteralPrefixLengthForOddCloser(text, pos, marker, remaining, canOpen, canClose);

                if (canClose && !preferInnerBold) {
                    while (remaining > 0) {
                        var closingIndex = pos + (runLen - remaining);
                        if (!TryCloseFrame(stack, marker, remaining, sourceMap, closingIndex, out int consumedClose)) break;
                        remaining -= consumedClose;
                    }
                }

                if (canOpen) {
                    if (splitDoubleRunIntoDualItalic || splitDoubleRunIntoRootDualItalic) {
                        stack.Push(new InlineFrame(FrameKind.Italic, marker, 1, new InlineSequence { AutoSpacing = false }, pos));
                        stack.Push(new InlineFrame(FrameKind.Italic, marker, 1, new InlineSequence { AutoSpacing = false }, pos + 1));
                        remaining -= 2;
                    }
                    else if (preferInnerBold) {
                        stack.Push(new InlineFrame(FrameKind.Bold, marker, 2, new InlineSequence { AutoSpacing = false }, pos));
                        remaining -= 2;
                    }

                    if (splitDoubleUnderscoreOpener && !splitDoubleRunIntoDualItalic && !splitDoubleRunIntoRootDualItalic) {
                        AddTextNode("_", pos, 1);
                        stack.Push(new InlineFrame(FrameKind.Italic, marker, 1, new InlineSequence { AutoSpacing = false }, pos + 1));
                        remaining -= 2;
                    }
                    else if (literalPrefixForOddCloser > 0 && !splitDoubleRunIntoDualItalic && !splitDoubleRunIntoRootDualItalic) {
                        AddTextNode(new string(marker, literalPrefixForOddCloser), pos, literalPrefixForOddCloser);
                        remaining -= literalPrefixForOddCloser;
                    }

                    while (remaining > 0) {
                        if (marker == '~') {
                            if (options.Subscript && !options.SingleTildeStrikethrough && remaining == 1) {
                                stack.Push(new InlineFrame(FrameKind.Subscript, marker, 1, new InlineSequence { AutoSpacing = false }, pos + (runLen - remaining)));
                                remaining -= 1;
                                continue;
                            }

                            int strikeDelimiterLength = remaining >= 2 ? 2 : (options.SingleTildeStrikethrough ? 1 : 2);
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

                        if (marker == '+') {
                            if (remaining >= 2) {
                                stack.Push(new InlineFrame(FrameKind.Inserted, marker, 2, new InlineSequence { AutoSpacing = false }, pos + (runLen - remaining)));
                                remaining -= 2;
                                continue;
                            }
                            break;
                        }

                        if (marker == '^') {
                            stack.Push(new InlineFrame(FrameKind.Superscript, marker, 1, new InlineSequence { AutoSpacing = false }, pos + (runLen - remaining)));
                            remaining -= 1;
                            continue;
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
                if (TryParseSupportedInlineHtmlTag(text, pos, options, state, sourceMap, allowLinks, allowImages, out int consumedHtmlTag, out var htmlNode)) {
                    AddRawNode(htmlNode, pos, consumedHtmlTag);
                    pos += consumedHtmlTag;
                    continue;
                }

                if (TryConsumeRawInlineHtmlTag(text, pos, sourceMap, out int consumedRawHtmlTag, out string rawInlineHtml)) {
                    AddRawNode(new HtmlRawInline(rawInlineHtml), pos, consumedRawHtmlTag);
                    pos += consumedRawHtmlTag;
                    continue;
                }
            }

            // Footnote ref [^id]
            if (options.Footnotes && text[pos] == '[' && pos + 2 < text.Length && text[pos + 1] == '^') {
                int rb = text.IndexOf(']', pos + 2);
                if (rb > pos + 2) { var lab = text.Substring(pos + 2, rb - (pos + 2)); AddFootnoteRefNode(lab, pos, rb + 1 - pos); pos = rb + 1; continue; }
            }

            int start = pos; pos++;
            while (pos < text.Length && !IsPotentialInlineStart(text[pos], options.InlineHtml, allowLinks, allowImages)) {
                // Ensure our explicit inline handlers see these characters.
                if (text[pos] == '\n') break;
                if (text[pos] == '\\' && pos + 1 < text.Length && IsBackslashEscapable(text[pos + 1])) break;
                if (text[pos] == '&' && TryConsumeHtmlEntityText(text, pos, out _, out _)) break;
                if (text[pos] == '<' && IsAngleAutolinkStart(text, pos)) break;
                if (options.AutolinkUrls && (text[pos] == 'h' || text[pos] == 'H') && StartsWithHttp(text, pos, options, out _)) break;
                if (options.AutolinkWwwUrls && (text[pos] == 'w' || text[pos] == 'W') && StartsWithWww(text, pos, options, out _)) break;
                if (options.AutolinkBareSchemeUrls && IsBareSchemeAutolinkStartCandidate(text[pos]) && TryConsumeBareSchemeAutolink(text, pos, options, out _, out _, out _)) break;
                if (options.AutolinkEmails && IsEmailStartChar(text[pos]) && TryConsumePlainEmail(text, pos, options, out _, out _)) break;
                if (options.Abbreviations && allowLinks && TryConsumeAbbreviation(text, pos, state, out _)) break;
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

}
