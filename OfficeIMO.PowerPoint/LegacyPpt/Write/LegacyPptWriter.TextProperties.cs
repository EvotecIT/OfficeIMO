using OfficeIMO.Core.Internal.OpenXml;
using static OfficeIMO.PowerPoint.LegacyPpt.Internal.LegacyPptXml;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        // Character attributes the binary character style leaves to the TextSIException, and the hyperlink children the
        // interaction records carry, so they are not formatting of the run itself.
        private static readonly HashSet<string> CharacterMetadataAttributes = new(StringComparer.Ordinal) {
            "lang", "altLang", "noProof", "dirty", "err", "smtClean"
        };

        // Paragraph attributes the shared text ruler carries instead of the paragraph exception.
        private static readonly HashSet<string> ParagraphRulerAttributes = new(StringComparer.Ordinal) {
            "lvl", "marL", "indent", "defTabSz"
        };

        /// <summary>
        /// Writes a paragraph exception for <c>a:pPr</c> or a list-style level. With <paramref name="excludeRuler"/> the
        /// level, margins, default tab size and tab stops are left out, as the shared text ruler carries them.
        /// </summary>
        private static bool TryWriteParagraphException(Stream output,
            OfficeXmlNode properties,
            LegacyPptWriterFontCatalog fonts, out string? reason,
            bool allowAutoNumbering = false, bool excludeRuler = false) {
            reason = null;
            if (properties.IsNone) {
                WriteUInt32(output, 0);
                return true;
            }
            if (!HasOnlyAttributes(properties, excludeRuler ? ParagraphRulerAttributes : null, "algn", "marL", "indent",
                    "defTabSz", "fontAlgn", "rtl", "eaLnBrk",
                    "latinLnBrk", "hangingPunct")) {
                reason = "A master paragraph style uses attributes outside the base binary text-property contract.";
                return false;
            }
            if (!HasAtMostOneOfEachParagraphChild(properties, excludeRuler)) {
                reason = "A master paragraph style contains duplicate or unsupported formatting children.";
                return false;
            }
            OfficeXmlNode autoNumberedBullet = properties.Child(A, "buAutoNum");
            OfficeXmlNode pictureBullet = properties.Child(A, "buBlip");
            if (!autoNumberedBullet.IsNone && !allowAutoNumbering) {
                reason = "A master paragraph style uses automatic numbering that requires a TextMasterStyle9Atom.";
                return false;
            }

            uint masks = 0;
            ushort bulletFlags = 0;
            bool hasBulletFlags = false;
            OfficeXmlNode noBullet = properties.Child(A, "buNone");
            OfficeXmlNode characterBullet = properties.Child(A, "buChar");
            if (!noBullet.IsNone
                && (HasAnyAttribute(noBullet) || HasChildElements(noBullet))) {
                reason = "A master no-bullet marker contains unsupported metadata.";
                return false;
            }
            if (!noBullet.IsNone || !characterBullet.IsNone
                || !autoNumberedBullet.IsNone || !pictureBullet.IsNone) {
                masks |= 1U;
                hasBulletFlags = true;
                if (!characterBullet.IsNone || !autoNumberedBullet.IsNone
                    || !pictureBullet.IsNone) {
                    bulletFlags |= 1;
                }
            }
            OfficeXmlNode bulletFont = properties.Child(A, "buFont");
            OfficeXmlNode bulletFontText = properties.Child(A, "buFontTx");
            if (!bulletFontText.IsNone
                && (HasAnyAttribute(bulletFontText) || HasChildElements(bulletFontText))) {
                reason = "A master text-inherited bullet font contains unsupported metadata.";
                return false;
            }
            ushort? bulletFontIndex = null;
            if (!bulletFont.IsNone || !bulletFontText.IsNone) {
                masks |= 1U << 1;
                hasBulletFlags = true;
                if (!bulletFont.IsNone) {
                    if (!HasOnlyAttributes(bulletFont, "typeface")
                        || HasChildElements(bulletFont)) {
                        reason = "A master bullet font contains metadata that has no base binary equivalent.";
                        return false;
                    }
                    bulletFlags |= 1 << 1;
                    if (!fonts.TryGetOrAdd(bulletFont.Attribute("typeface"),
                            out ushort index, out reason)) return false;
                    bulletFontIndex = index;
                    masks |= 1U << 4;
                }
            }
            OfficeXmlNode bulletColor = properties.Child(A, "buClr");
            OfficeXmlNode bulletColorText = properties.Child(A, "buClrTx");
            if (!bulletColorText.IsNone
                && (HasAnyAttribute(bulletColorText) || HasChildElements(bulletColorText))) {
                reason = "A master text-inherited bullet color contains unsupported metadata.";
                return false;
            }
            byte[]? bulletColorBytes = null;
            if (!bulletColor.IsNone || !bulletColorText.IsNone) {
                masks |= 1U << 2;
                hasBulletFlags = true;
                if (!bulletColor.IsNone) {
                    bulletFlags |= 1 << 2;
                    if (!TryReadBinaryColor(bulletColor, out bulletColorBytes,
                            out reason)) return false;
                    masks |= 1U << 5;
                }
            }
            OfficeXmlNode bulletPercent = properties.Child(A, "buSzPct");
            OfficeXmlNode bulletPoints = properties.Child(A, "buSzPts");
            OfficeXmlNode bulletSizeText = properties.Child(A, "buSzTx");
            if (!bulletSizeText.IsNone
                && (HasAnyAttribute(bulletSizeText) || HasChildElements(bulletSizeText))) {
                reason = "A master text-inherited bullet size contains unsupported metadata.";
                return false;
            }
            short? bulletSize = null;
            if (!bulletPercent.IsNone || !bulletPoints.IsNone
                || !bulletSizeText.IsNone) {
                masks |= 1U << 3;
                hasBulletFlags = true;
                if (!bulletPercent.IsNone || !bulletPoints.IsNone) {
                    OfficeXmlNode numericSize = !bulletPercent.IsNone ? bulletPercent : bulletPoints;
                    if (!HasOnlyAttributes(numericSize, "val")
                        || HasChildElements(numericSize)) {
                        reason = "A master bullet size contains metadata that has no base binary equivalent.";
                        return false;
                    }
                    bulletFlags |= 1 << 3;
                    int? percentValue = bulletPercent.IsNone ? null : bulletPercent.Int("val");
                    int? pointValue = bulletPoints.IsNone ? null : bulletPoints.Int("val");
                    if (percentValue.HasValue) {
                        int value = percentValue.Value;
                        if (value % 1000 != 0 || value < 25000
                            || value > 400000) {
                            reason = "A master bullet percentage cannot be represented exactly by the binary PowerPoint integer percentage field.";
                            return false;
                        }
                        bulletSize = checked((short)(value / 1000));
                    } else if (pointValue.HasValue) {
                        int value = pointValue.Value;
                        if (value % 100 != 0 || value < 100
                            || value > 400000) {
                            reason = "A master bullet point size cannot be represented exactly by the binary PowerPoint point field.";
                            return false;
                        }
                        bulletSize = checked((short)-(value / 100));
                    } else {
                        reason = "A master bullet size has no numeric value.";
                        return false;
                    }
                    masks |= 1U << 6;
                }
            }
            char? bulletCharacter = null;
            if (!characterBullet.IsNone) {
                if (!HasOnlyAttributes(characterBullet, "char")
                    || HasChildElements(characterBullet)) {
                    reason = "A master character bullet contains metadata that has no base binary equivalent.";
                    return false;
                }
                string value = characterBullet.Attribute("char") ?? string.Empty;
                if (value.Length != 1 || value[0] == '\0') {
                    reason = "A binary PowerPoint character bullet must contain exactly one non-NUL UTF-16 character.";
                    return false;
                }
                bulletCharacter = value[0];
                masks |= 1U << 7;
            }

            short? leftMargin = null;
            int? leftMarginValue = excludeRuler ? null : properties.Int("marL");
            if (leftMarginValue.HasValue) {
                if (!TryToMasterInt16(leftMarginValue.Value,
                        out short value)) {
                    reason = "A master paragraph left margin lies outside the binary PowerPoint range.";
                    return false;
                }
                leftMargin = value;
                masks |= 1U << 8;
            }
            short? indent = null;
            int? indentValue = excludeRuler ? null : properties.Int("indent");
            if (indentValue.HasValue) {
                if (!TryToMasterInt16(indentValue.Value,
                        out short value)) {
                    reason = "A master paragraph first-line indent lies outside the binary PowerPoint range.";
                    return false;
                }
                indent = value;
                masks |= 1U << 10;
            }
            ushort? alignment = null;
            string? alignmentToken = properties.Attribute("algn");
            if (IsParagraphAlignmentToken(alignmentToken)) {
                if (!TryMapAlignment(alignmentToken, out ushort value)) {
                    reason = "A master paragraph alignment has no base binary PowerPoint equivalent.";
                    return false;
                }
                alignment = value;
                masks |= 1U << 11;
            }
            if (!TryReadSpacing(properties.Child(A, "lnSpc"),
                    out short? lineSpacing, out reason)
                || !TryReadSpacing(properties.Child(A, "spcBef"),
                    out short? spaceBefore, out reason)
                || !TryReadSpacing(properties.Child(A, "spcAft"),
                    out short? spaceAfter, out reason)) return false;
            if (lineSpacing.HasValue) masks |= 1U << 12;
            if (spaceBefore.HasValue) masks |= 1U << 13;
            if (spaceAfter.HasValue) masks |= 1U << 14;
            short? defaultTab = null;
            int? defaultTabValue = excludeRuler ? null : properties.Int("defTabSz");
            if (defaultTabValue.HasValue) {
                if (!TryToMasterInt16(defaultTabValue.Value,
                        out short value) || value < 0) {
                    reason = "A master default tab size lies outside the binary PowerPoint range.";
                    return false;
                }
                defaultTab = value;
                masks |= 1U << 15;
            }
            ushort? fontAlignment = null;
            string? fontAlignmentToken = properties.Attribute("fontAlgn");
            if (fontAlignmentToken is "auto" or "t" or "ctr" or "base" or "b") {
                if (!TryMapFontAlignment(fontAlignmentToken, out ushort value)) {
                    reason = "A master font alignment has no base binary PowerPoint equivalent.";
                    return false;
                }
                fontAlignment = value;
                masks |= 1U << 16;
            }
            ushort wrapFlags = 0;
            bool hasWrapFlags = false;
            bool? eastAsianLineBreak = properties.Bool("eaLnBrk");
            if (eastAsianLineBreak.HasValue) {
                masks |= 1U << 17;
                hasWrapFlags = true;
                if (eastAsianLineBreak.Value) wrapFlags |= 1;
            }
            bool? latinLineBreak = properties.Bool("latinLnBrk");
            if (latinLineBreak.HasValue) {
                masks |= 1U << 18;
                hasWrapFlags = true;
                if (!latinLineBreak.Value) wrapFlags |= 1 << 1;
            }
            bool? hangingPunctuation = properties.Bool("hangingPunct");
            if (hangingPunctuation.HasValue) {
                masks |= 1U << 19;
                hasWrapFlags = true;
                if (hangingPunctuation.Value) wrapFlags |= 1 << 2;
            }
            OfficeXmlNode tabList = excludeRuler ? default : properties.Child(A, "tabLst");
            var tabStops = new List<KeyValuePair<short, ushort>>();
            if (!tabList.IsNone) {
                if (!TryReadTabStops(tabList, tabStops, out bool unsupportedList)) {
                    reason = unsupportedList
                        ? "A master tab-stop list contains unsupported content."
                        : "A master tab stop has an unsupported position or alignment.";
                    return false;
                }
                if (tabStops.Count > ushort.MaxValue) {
                    reason = "A master paragraph has too many binary tab stops.";
                    return false;
                }
                masks |= 1U << 20;
            }
            ushort? direction = null;
            bool? rightToLeft = properties.Bool("rtl");
            if (rightToLeft.HasValue) {
                direction = rightToLeft.Value ? (ushort)1 : (ushort)0;
                masks |= 1U << 21;
            }

            WriteUInt32(output, masks);
            if (hasBulletFlags) WriteUInt16(output, bulletFlags);
            if (bulletCharacter.HasValue) WriteUInt16(output,
                bulletCharacter.Value);
            if (bulletFontIndex.HasValue) WriteUInt16(output,
                bulletFontIndex.Value);
            if (bulletSize.HasValue) WriteInt16(output, bulletSize.Value);
            if (bulletColorBytes != null) output.Write(bulletColorBytes, 0,
                bulletColorBytes.Length);
            if (alignment.HasValue) WriteUInt16(output, alignment.Value);
            if (lineSpacing.HasValue) WriteInt16(output, lineSpacing.Value);
            if (spaceBefore.HasValue) WriteInt16(output, spaceBefore.Value);
            if (spaceAfter.HasValue) WriteInt16(output, spaceAfter.Value);
            if (leftMargin.HasValue) WriteInt16(output, leftMargin.Value);
            if (indent.HasValue) WriteInt16(output, indent.Value);
            if (defaultTab.HasValue) WriteInt16(output, defaultTab.Value);
            if (!tabList.IsNone) {
                WriteUInt16(output, checked((ushort)tabStops.Count));
                foreach (KeyValuePair<short, ushort> tab in tabStops) {
                    WriteInt16(output, tab.Key);
                    WriteUInt16(output, tab.Value);
                }
            }
            if (fontAlignment.HasValue) WriteUInt16(output,
                fontAlignment.Value);
            if (hasWrapFlags) WriteUInt16(output, wrapFlags);
            if (direction.HasValue) WriteUInt16(output, direction.Value);
            return true;
        }

        /// <summary>
        /// Writes a character exception for run properties (<c>a:rPr</c>, <c>a:endParaRPr</c>, <c>a:defRPr</c>). With
        /// <paramref name="excludeMetadata"/> the proofing and language attributes and the hyperlink children are left
        /// out, as other records carry them.
        /// </summary>
        private static bool TryWriteCharacterException(Stream output,
            OfficeXmlNode properties,
            LegacyPptWriterFontCatalog fonts, out string? reason,
            byte? ppt9RunId = null, bool excludeMetadata = false) {
            reason = null;
            if (properties.IsNone && !ppt9RunId.HasValue) {
                WriteUInt32(output, 0);
                return true;
            }
            if (!properties.IsNone && !HasOnlyAttributes(properties,
                    excludeMetadata ? CharacterMetadataAttributes : null, "b", "i", "u", "kumimoji",
                    "sz", "baseline")) {
                reason = "A text run uses attributes outside the base binary character-style contract.";
                return false;
            }
            foreach (OfficeXmlNode child in properties.Children()) {
                if (excludeMetadata && IsHyperlinkChild(child)) continue;
                if (!child.Is(A, "solidFill") && !child.Is(A, "latin")
                    && !child.Is(A, "ea") && !child.Is(A, "sym")) {
                    reason = $"A text run contains unsupported '{child.LocalName}' content.";
                    return false;
                }
            }
            if (!properties.IsNone
                && (properties.Children(A, "solidFill").Count > 1
                    || properties.Children(A, "latin").Count > 1
                    || properties.Children(A, "ea").Count > 1
                    || properties.Children(A, "sym").Count > 1)) {
                reason = "A text run contains duplicate color or typeface elements.";
                return false;
            }
            uint masks = 0;
            ushort style = 0;
            bool hasStyle = false;
            AddStyleFlag(properties.IsNone ? null : properties.Bool("b"), 0, ref masks, ref style,
                ref hasStyle);
            AddStyleFlag(properties.IsNone ? null : properties.Bool("i"), 1, ref masks, ref style,
                ref hasStyle);
            string? underline = properties.IsNone ? null : properties.Attribute("u");
            if (IsUnderlineToken(underline)) {
                if (underline != "none" && underline != "sng") {
                    reason = "Only none and single underline map exactly to the base binary character style.";
                    return false;
                }
                masks |= 1U << 2;
                hasStyle = true;
                if (underline == "sng") style |= 1 << 2;
            }
            AddStyleFlag(properties.IsNone ? null : properties.Bool("kumimoji"), 7, ref masks, ref style,
                ref hasStyle);
            if (ppt9RunId.HasValue) {
                if (ppt9RunId.Value > 0x0F) {
                    reason = "A PPT9 text run identifier is outside the four-bit range.";
                    return false;
                }
                masks |= 0x00003C00U;
                style |= checked((ushort)(ppt9RunId.Value << 10));
                hasStyle = true;
            }

            ushort? latinFont = null;
            OfficeXmlNode latin = properties.Child(A, "latin");
            if (!latin.IsNone) {
                if (!HasOnlyAttributes(latin, "typeface")
                    || HasChildElements(latin)
                    || !fonts.TryGetOrAdd(latin.Attribute("typeface"),
                        out ushort index, out reason)) return false;
                latinFont = index;
                masks |= 1U << 16;
            }
            ushort? eastAsianFont = null;
            OfficeXmlNode eastAsian = properties.Child(A, "ea");
            if (!eastAsian.IsNone) {
                if (!HasOnlyAttributes(eastAsian, "typeface")
                    || HasChildElements(eastAsian)
                    || !fonts.TryGetOrAdd(eastAsian.Attribute("typeface"),
                        out ushort index, out reason)) return false;
                eastAsianFont = index;
                masks |= 1U << 21;
            }
            ushort? symbolFont = null;
            OfficeXmlNode symbol = properties.Child(A, "sym");
            if (!symbol.IsNone) {
                if (!HasOnlyAttributes(symbol, "typeface")
                    || HasChildElements(symbol)
                    || !fonts.TryGetOrAdd(symbol.Attribute("typeface"),
                        out ushort index, out reason)) return false;
                symbolFont = index;
                masks |= 1U << 23;
            }
            short? fontSize = null;
            int? fontSizeValue = properties.IsNone ? null : properties.Int("sz");
            if (fontSizeValue.HasValue) {
                int value = fontSizeValue.Value;
                if (value % 100 != 0 || value < 100 || value > 400000) {
                    reason = "A master font size cannot be represented exactly as whole binary PowerPoint points.";
                    return false;
                }
                fontSize = checked((short)(value / 100));
                masks |= 1U << 17;
            }
            byte[]? color = null;
            OfficeXmlNode fill = properties.Child(A, "solidFill");
            if (!fill.IsNone) {
                if (!TryReadBinaryColor(fill, out color, out reason)) {
                    return false;
                }
                masks |= 1U << 18;
            }
            short? baseline = null;
            int? baselineValue = properties.IsNone ? null : properties.Int("baseline");
            if (baselineValue.HasValue) {
                int value = baselineValue.Value;
                if (value % 1000 != 0 || value < -100000
                    || value > 100000) {
                    reason = "A master baseline offset cannot be represented exactly as a binary PowerPoint percentage.";
                    return false;
                }
                baseline = checked((short)(value / 1000));
                masks |= 1U << 19;
            }

            WriteUInt32(output, masks);
            if (hasStyle) WriteUInt16(output, style);
            if (latinFont.HasValue) WriteUInt16(output, latinFont.Value);
            if (eastAsianFont.HasValue) WriteUInt16(output,
                eastAsianFont.Value);
            if (symbolFont.HasValue) WriteUInt16(output, symbolFont.Value);
            if (fontSize.HasValue) WriteInt16(output, fontSize.Value);
            if (color != null) output.Write(color, 0, color.Length);
            if (baseline.HasValue) WriteInt16(output, baseline.Value);
            return true;
        }

        private static void AddStyleFlag(bool? value, int bit,
            ref uint masks, ref ushort style, ref bool hasStyle) {
            if (!value.HasValue) return;
            masks |= 1U << bit;
            hasStyle = true;
            if (value.Value) style |= checked((ushort)(1 << bit));
        }

        private static bool IsHyperlinkChild(OfficeXmlNode child) =>
            child.Is(A, "hlinkClick") || child.Is(A, "hlinkMouseOver");

        /// <summary>
        /// Whether run properties carry formatting beyond the proofing, language and hyperlink metadata other records
        /// hold.
        /// </summary>
        private static bool HasCharacterFormatting(OfficeXmlNode properties) {
            if (properties.IsNone) return false;
            foreach (string name in AttributeLocalNames(properties)) {
                if (!CharacterMetadataAttributes.Contains(name)) return true;
            }
            foreach (OfficeXmlNode child in properties.Children()) {
                if (!IsHyperlinkChild(child)) return true;
            }
            return false;
        }

        /// <summary>Whether paragraph properties carry formatting beyond what the shared text ruler holds.</summary>
        private static bool HasParagraphFormattingOutsideRuler(OfficeXmlNode properties) {
            if (properties.IsNone) return false;
            foreach (string name in AttributeLocalNames(properties)) {
                if (!ParagraphRulerAttributes.Contains(name)) return true;
            }
            foreach (OfficeXmlNode child in properties.Children()) {
                if (!child.Is(A, "tabLst")) return true;
            }
            return false;
        }

        private static readonly string[] ParagraphChildNames = {
            "lnSpc", "spcBef", "spcAft", "buClrTx", "buClr", "buSzTx", "buSzPct", "buSzPts", "buFontTx", "buFont",
            "buNone", "buChar", "buAutoNum", "buBlip", "tabLst"
        };

        private static bool HasAtMostOneOfEachParagraphChild(
            OfficeXmlNode properties, bool excludeRuler = false) {
            if (properties.Children(A, "defRPr").Count > 1) {
                return false;
            }
            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (OfficeXmlNode child in properties.Children()) {
                if (child.Is(A, "defRPr")) continue;
                if (excludeRuler && child.Is(A, "tabLst")) continue;
                if (child.NamespaceId != A
                    || Array.IndexOf(ParagraphChildNames, child.LocalName) < 0) return false;
                if (!seen.Add(child.LocalName)) return false;
            }
            return CountPresent(properties, "buNone", "buChar", "buAutoNum", "buBlip") <= 1
                && CountPresent(properties, "buFont", "buFontTx") <= 1
                && CountPresent(properties, "buClr", "buClrTx") <= 1
                && CountPresent(properties, "buSzPct", "buSzPts", "buSzTx") <= 1;
        }

        private static int CountPresent(OfficeXmlNode parent, params string[] localNames) =>
            localNames.Count(name => !parent.Child(A, name).IsNone);

        private static bool HasAnyAttribute(OfficeXmlNode element) => AttributeLocalNames(element).Count != 0;

        private static bool HasOnlyAttributes(OfficeXmlNode element,
            params string[] allowed) => HasOnlyAttributes(element, null, allowed);

        private static bool HasOnlyAttributes(OfficeXmlNode element,
            HashSet<string>? ignored, params string[] allowed) {
            foreach (string name in AttributeLocalNames(element)) {
                if (ignored != null && ignored.Contains(name)) continue;
                if (Array.IndexOf(allowed, name) < 0) return false;
            }
            return true;
        }

        /// <summary>Reads <c>a:tabLst</c> into binary tab stops; false for anything without an exact mapping.</summary>
        private static bool TryReadTabStops(OfficeXmlNode list, List<KeyValuePair<short, ushort>> result,
            out bool unsupportedList) {
            unsupportedList = HasAnyAttribute(list) || list.Children(A, "tab").Count
                != list.Children().Count;
            if (unsupportedList) return false;
            foreach (OfficeXmlNode tab in list.Children(A, "tab")) {
                int? position = tab.Int("pos");
                string? alignment = tab.Attribute("algn");
                if (!HasOnlyAttributes(tab, "pos", "algn")
                    || !position.HasValue
                    || !TryToMasterInt16(position.Value,
                        out short binaryPosition)
                    || !TryMapTabAlignment(alignment,
                        out ushort binaryAlignment)) {
                    return false;
                }
                result.Add(new KeyValuePair<short, ushort>(binaryPosition,
                    binaryAlignment));
            }
            return true;
        }

        private static bool TryReadSpacing(OfficeXmlNode spacing,
            out short? value, out string? reason) {
            value = null;
            reason = null;
            if (spacing.IsNone) return true;
            if (HasAnyAttribute(spacing) || spacing.Children().Count != 1) {
                reason = "A master paragraph spacing element contains unsupported content.";
                return false;
            }
            OfficeXmlNode percent = spacing.Child(A, "spcPct");
            if (!percent.IsNone && percent.Int("val") is int rawPercent) {
                if (!HasOnlyAttributes(percent, "val")
                    || HasChildElements(percent)) {
                    reason = "A master paragraph percentage spacing contains unsupported metadata.";
                    return false;
                }
                if (rawPercent % 1000 != 0 || rawPercent < 0 || rawPercent > 13200000) {
                    reason = "A master paragraph percentage spacing cannot be represented exactly by the binary field.";
                    return false;
                }
                value = checked((short)(rawPercent / 1000));
                return true;
            }
            OfficeXmlNode points = spacing.Child(A, "spcPts");
            if (!points.IsNone && points.Int("val") is int raw) {
                if (!HasOnlyAttributes(points, "val")
                    || HasChildElements(points)) {
                    reason = "A master paragraph point spacing contains unsupported metadata.";
                    return false;
                }
                if (raw < 0) {
                    reason = "A master paragraph point spacing cannot be represented exactly by the binary 1/8-point field.";
                    return false;
                }
                long binary = -(long)Math.Round(raw / 12.5D,
                    MidpointRounding.AwayFromZero);
                if (binary < short.MinValue || binary > -1) {
                    reason = "A master paragraph point spacing lies outside the binary field range.";
                    return false;
                }
                long projected = (long)Math.Round(-binary * 12.5D,
                    MidpointRounding.AwayFromZero);
                if (projected != raw) {
                    reason = "A master paragraph point spacing cannot be represented exactly by the binary 1/8-point field.";
                    return false;
                }
                value = checked((short)binary);
                return true;
            }
            reason = "A master paragraph spacing element has no supported numeric child.";
            return false;
        }

        private static bool TryReadBinaryColor(OfficeXmlNode parent,
            out byte[]? color, out string? reason) {
            color = null;
            reason = null;
            if (HasAnyAttribute(parent) || parent.Children().Count != 1) {
                reason = "A master text color contains transforms or unsupported color metadata.";
                return false;
            }
            OfficeXmlNode rgb = parent.Child(A, "srgbClr");
            if (!rgb.IsNone && rgb.Attribute("val") is string value) {
                string normalized = value.Trim().TrimStart('#');
                if (normalized.Length != 6 || !normalized.All(Uri.IsHexDigit)
                    || !HasOnlyAttributes(rgb, "val")
                    || HasChildElements(rgb)) {
                    reason = "A master text RGB color is not an exact six-digit value without transforms.";
                    return false;
                }
                color = new[] {
                    Convert.ToByte(normalized.Substring(0, 2), 16),
                    Convert.ToByte(normalized.Substring(2, 2), 16),
                    Convert.ToByte(normalized.Substring(4, 2), 16),
                    (byte)0xFE
                };
                return true;
            }
            OfficeXmlNode scheme = parent.Child(A, "schemeClr");
            if (!scheme.IsNone
                && HasOnlyAttributes(scheme, "val")
                && !HasChildElements(scheme)
                && TryMapSchemeColor(scheme.Attribute("val"), out byte index)) {
                color = new byte[] { 0, 0, 0, index };
                return true;
            }
            reason = "A master text color has no base binary RGB or eight-slot scheme equivalent.";
            return false;
        }

        private static bool TryMapSchemeColor(string? token,
            out byte index) {
            switch (token) {
                case "bg1":
                case "lt1":
                    index = 0;
                    return true;
                case "tx1":
                case "dk1":
                    index = 1;
                    return true;
                case "accent4":
                    index = 2;
                    return true;
                case "tx2":
                case "dk2":
                    index = 3;
                    return true;
                case "bg2":
                case "lt2":
                    index = 4;
                    return true;
                case "accent1":
                    index = 5;
                    return true;
                case "accent2":
                    index = 6;
                    return true;
                case "accent3":
                    index = 7;
                    return true;
                default:
                    index = byte.MaxValue;
                    return false;
            }
        }

        private static bool TryToMasterInt16(long emus, out short value) {
            long converted = checked((long)Math.Round(emus / 1587.5D,
                MidpointRounding.AwayFromZero));
            if (converted < short.MinValue || converted > short.MaxValue) {
                value = 0;
                return false;
            }
            value = checked((short)converted);
            long projected = checked((long)Math.Round(value * 1587.5D,
                MidpointRounding.AwayFromZero));
            return projected == emus;
        }

        // ST_TextAlignType: every token has a binary alignment.
        private static bool IsParagraphAlignmentToken(string? token) =>
            token is "l" or "ctr" or "r" or "just" or "justLow" or "dist" or "thaiDist";

        private static bool TryMapAlignment(string? token,
            out ushort mapped) {
            switch (token) {
                case "l": mapped = 0; return true;
                case "ctr": mapped = 1; return true;
                case "r": mapped = 2; return true;
                case "just": mapped = 3; return true;
                case "dist": mapped = 4; return true;
                case "thaiDist": mapped = 5; return true;
                case "justLow": mapped = 6; return true;
                default:
                    mapped = ushort.MaxValue;
                    return false;
            }
        }

        private static bool TryMapFontAlignment(string? token,
            out ushort mapped) {
            switch (token) {
                case "base": mapped = 0; return true;
                case "t": mapped = 1; return true;
                case "ctr": mapped = 2; return true;
                case "b": mapped = 3; return true;
                default:
                    mapped = ushort.MaxValue;
                    return false;
            }
        }

        private static bool TryMapTabAlignment(string? token,
            out ushort mapped) {
            switch (token) {
                case "l": mapped = 0; return true;
                case "ctr": mapped = 1; return true;
                case "r": mapped = 2; return true;
                case "dec": mapped = 3; return true;
                default:
                    mapped = ushort.MaxValue;
                    return false;
            }
        }

        // ST_TextUnderlineType; a value outside it reads as absent, as the schema type would.
        private static bool IsUnderlineToken(string? token) => token is "none" or "words" or "sng" or "dbl" or "heavy"
            or "dotted" or "dottedHeavy" or "dash" or "dashHeavy" or "dashLong" or "dashLongHeavy" or "dotDash"
            or "dotDashHeavy" or "dotDotDash" or "dotDotDashHeavy" or "wavy" or "wavyHeavy" or "wavyDbl";
    }
}
