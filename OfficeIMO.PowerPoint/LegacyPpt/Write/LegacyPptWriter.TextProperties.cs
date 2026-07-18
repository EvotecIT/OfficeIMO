using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        private static bool TryWriteParagraphException(Stream output,
            A.TextParagraphPropertiesType? properties,
            LegacyPptWriterFontCatalog fonts, out string? reason,
            bool allowAutoNumbering = false) {
            reason = null;
            if (properties == null) {
                WriteUInt32(output, 0);
                return true;
            }
            if (!HasOnlyAttributes(properties, "algn", "marL", "indent",
                    "defTabSz", "fontAlgn", "rtl", "eaLnBrk",
                    "latinLnBrk", "hangingPunct")) {
                reason = "A master paragraph style uses attributes outside the base binary text-property contract.";
                return false;
            }
            if (!HasAtMostOneOfEachParagraphChild(properties)) {
                reason = "A master paragraph style contains duplicate or unsupported formatting children.";
                return false;
            }
            A.AutoNumberedBullet? autoNumberedBullet = properties
                .GetFirstChild<A.AutoNumberedBullet>();
            A.PictureBullet? pictureBullet = properties
                .GetFirstChild<A.PictureBullet>();
            if (autoNumberedBullet != null && !allowAutoNumbering) {
                reason = "A master paragraph style uses automatic numbering that requires a TextMasterStyle9Atom.";
                return false;
            }

            uint masks = 0;
            ushort bulletFlags = 0;
            bool hasBulletFlags = false;
            A.NoBullet? noBullet = properties.GetFirstChild<A.NoBullet>();
            A.CharacterBullet? characterBullet = properties
                .GetFirstChild<A.CharacterBullet>();
            if (noBullet != null
                && (noBullet.HasAttributes
                    || noBullet.ChildElements.Count != 0)) {
                reason = "A master no-bullet marker contains unsupported metadata.";
                return false;
            }
            if (noBullet != null || characterBullet != null
                || autoNumberedBullet != null || pictureBullet != null) {
                masks |= 1U;
                hasBulletFlags = true;
                if (characterBullet != null || autoNumberedBullet != null
                    || pictureBullet != null) {
                    bulletFlags |= 1;
                }
            }
            A.BulletFont? bulletFont = properties.GetFirstChild<A.BulletFont>();
            A.BulletFontText? bulletFontText = properties
                .GetFirstChild<A.BulletFontText>();
            if (bulletFontText != null
                && (bulletFontText.HasAttributes
                    || bulletFontText.ChildElements.Count != 0)) {
                reason = "A master text-inherited bullet font contains unsupported metadata.";
                return false;
            }
            ushort? bulletFontIndex = null;
            if (bulletFont != null || bulletFontText != null) {
                masks |= 1U << 1;
                hasBulletFlags = true;
                if (bulletFont != null) {
                    if (!HasOnlyAttributes(bulletFont, "typeface")
                        || bulletFont.ChildElements.Count != 0) {
                        reason = "A master bullet font contains metadata that has no base binary equivalent.";
                        return false;
                    }
                    bulletFlags |= 1 << 1;
                    if (!fonts.TryGetOrAdd(bulletFont.Typeface?.Value,
                            out ushort index, out reason)) return false;
                    bulletFontIndex = index;
                    masks |= 1U << 4;
                }
            }
            A.BulletColor? bulletColor = properties.GetFirstChild<A.BulletColor>();
            A.BulletColorText? bulletColorText = properties
                .GetFirstChild<A.BulletColorText>();
            if (bulletColorText != null
                && (bulletColorText.HasAttributes
                    || bulletColorText.ChildElements.Count != 0)) {
                reason = "A master text-inherited bullet color contains unsupported metadata.";
                return false;
            }
            byte[]? bulletColorBytes = null;
            if (bulletColor != null || bulletColorText != null) {
                masks |= 1U << 2;
                hasBulletFlags = true;
                if (bulletColor != null) {
                    bulletFlags |= 1 << 2;
                    if (!TryReadBinaryColor(bulletColor, out bulletColorBytes,
                            out reason)) return false;
                    masks |= 1U << 5;
                }
            }
            A.BulletSizePercentage? bulletPercent = properties
                .GetFirstChild<A.BulletSizePercentage>();
            A.BulletSizePoints? bulletPoints = properties
                .GetFirstChild<A.BulletSizePoints>();
            A.BulletSizeText? bulletSizeText = properties
                .GetFirstChild<A.BulletSizeText>();
            if (bulletSizeText != null
                && (bulletSizeText.HasAttributes
                    || bulletSizeText.ChildElements.Count != 0)) {
                reason = "A master text-inherited bullet size contains unsupported metadata.";
                return false;
            }
            short? bulletSize = null;
            if (bulletPercent != null || bulletPoints != null
                || bulletSizeText != null) {
                masks |= 1U << 3;
                hasBulletFlags = true;
                if (bulletPercent != null || bulletPoints != null) {
                    OpenXmlElement numericSize = (OpenXmlElement?)bulletPercent
                        ?? bulletPoints!;
                    if (!HasOnlyAttributes(numericSize, "val")
                        || numericSize.ChildElements.Count != 0) {
                        reason = "A master bullet size contains metadata that has no base binary equivalent.";
                        return false;
                    }
                    bulletFlags |= 1 << 3;
                    if (bulletPercent?.Val?.HasValue == true) {
                        int value = bulletPercent.Val.Value;
                        if (value % 1000 != 0 || value < 25000
                            || value > 400000) {
                            reason = "A master bullet percentage cannot be represented exactly by the binary PowerPoint integer percentage field.";
                            return false;
                        }
                        bulletSize = checked((short)(value / 1000));
                    } else if (bulletPoints?.Val?.HasValue == true) {
                        int value = bulletPoints.Val.Value;
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
            if (characterBullet != null) {
                if (!HasOnlyAttributes(characterBullet, "char")
                    || characterBullet.ChildElements.Count != 0) {
                    reason = "A master character bullet contains metadata that has no base binary equivalent.";
                    return false;
                }
                string value = characterBullet.Char?.Value ?? string.Empty;
                if (value.Length != 1 || value[0] == '\0') {
                    reason = "A binary PowerPoint character bullet must contain exactly one non-NUL UTF-16 character.";
                    return false;
                }
                bulletCharacter = value[0];
                masks |= 1U << 7;
            }

            short? leftMargin = null;
            if (properties.LeftMargin?.HasValue == true) {
                if (!TryToMasterInt16(properties.LeftMargin.Value,
                        out short value)) {
                    reason = "A master paragraph left margin lies outside the binary PowerPoint range.";
                    return false;
                }
                leftMargin = value;
                masks |= 1U << 8;
            }
            short? indent = null;
            if (properties.Indent?.HasValue == true) {
                if (!TryToMasterInt16(properties.Indent.Value,
                        out short value)) {
                    reason = "A master paragraph first-line indent lies outside the binary PowerPoint range.";
                    return false;
                }
                indent = value;
                masks |= 1U << 10;
            }
            ushort? alignment = null;
            if (properties.Alignment?.HasValue == true) {
                if (!TryMapAlignment(properties.Alignment.Value,
                        out ushort value)) {
                    reason = "A master paragraph alignment has no base binary PowerPoint equivalent.";
                    return false;
                }
                alignment = value;
                masks |= 1U << 11;
            }
            if (!TryReadSpacing(properties.GetFirstChild<A.LineSpacing>(),
                    out short? lineSpacing, out reason)
                || !TryReadSpacing(properties.GetFirstChild<A.SpaceBefore>(),
                    out short? spaceBefore, out reason)
                || !TryReadSpacing(properties.GetFirstChild<A.SpaceAfter>(),
                    out short? spaceAfter, out reason)) return false;
            if (lineSpacing.HasValue) masks |= 1U << 12;
            if (spaceBefore.HasValue) masks |= 1U << 13;
            if (spaceAfter.HasValue) masks |= 1U << 14;
            short? defaultTab = null;
            if (properties.DefaultTabSize?.HasValue == true) {
                if (!TryToMasterInt16(properties.DefaultTabSize.Value,
                        out short value) || value < 0) {
                    reason = "A master default tab size lies outside the binary PowerPoint range.";
                    return false;
                }
                defaultTab = value;
                masks |= 1U << 15;
            }
            ushort? fontAlignment = null;
            if (properties.FontAlignment?.HasValue == true) {
                if (!TryMapFontAlignment(properties.FontAlignment.Value,
                        out ushort value)) {
                    reason = "A master font alignment has no base binary PowerPoint equivalent.";
                    return false;
                }
                fontAlignment = value;
                masks |= 1U << 16;
            }
            ushort wrapFlags = 0;
            bool hasWrapFlags = false;
            if (properties.EastAsianLineBreak?.HasValue == true) {
                masks |= 1U << 17;
                hasWrapFlags = true;
                if (properties.EastAsianLineBreak.Value) wrapFlags |= 1;
            }
            if (properties.LatinLineBreak?.HasValue == true) {
                masks |= 1U << 18;
                hasWrapFlags = true;
                if (!properties.LatinLineBreak.Value) wrapFlags |= 1 << 1;
            }
            if (properties.Height?.HasValue == true) {
                masks |= 1U << 19;
                hasWrapFlags = true;
                if (properties.Height.Value) wrapFlags |= 1 << 2;
            }
            A.TabStopList? tabList = properties.GetFirstChild<A.TabStopList>();
            var tabStops = new List<KeyValuePair<short, ushort>>();
            if (tabList != null) {
                if (tabList.HasAttributes || tabList.Elements<A.TabStop>().Count()
                        != tabList.ChildElements.Count) {
                    reason = "A master tab-stop list contains unsupported content.";
                    return false;
                }
                foreach (A.TabStop tab in tabList.Elements<A.TabStop>()) {
                    if (!HasOnlyAttributes(tab, "pos", "algn")
                        || tab.Position?.HasValue != true
                        || tab.Alignment?.HasValue != true
                        || !TryToMasterInt16(tab.Position.Value,
                            out short position)
                        || !TryMapTabAlignment(tab.Alignment.Value,
                            out ushort tabAlignment)) {
                        reason = "A master tab stop has an unsupported position or alignment.";
                        return false;
                    }
                    tabStops.Add(new KeyValuePair<short, ushort>(position,
                        tabAlignment));
                }
                if (tabStops.Count > ushort.MaxValue) {
                    reason = "A master paragraph has too many binary tab stops.";
                    return false;
                }
                masks |= 1U << 20;
            }
            ushort? direction = null;
            if (properties.RightToLeft?.HasValue == true) {
                direction = properties.RightToLeft.Value ? (ushort)1 : (ushort)0;
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
            if (tabList != null) {
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

        private static bool TryWriteCharacterException(Stream output,
            A.TextCharacterPropertiesType? properties,
            LegacyPptWriterFontCatalog fonts, out string? reason,
            byte? ppt9RunId = null) {
            reason = null;
            if (properties == null && !ppt9RunId.HasValue) {
                WriteUInt32(output, 0);
                return true;
            }
            if (properties != null && !HasOnlyAttributes(properties, "b", "i", "u", "kumimoji",
                    "sz", "baseline")) {
                reason = "A text run uses attributes outside the base binary character-style contract.";
                return false;
            }
            if (properties != null) {
                foreach (OpenXmlElement child in properties.ChildElements) {
                    if (child is not A.SolidFill and not A.LatinFont
                        and not A.EastAsianFont and not A.SymbolFont) {
                        reason = $"A text run contains unsupported '{child.LocalName}' content.";
                        return false;
                    }
                }
            }
            if (properties != null
                && (properties.Elements<A.SolidFill>().Count() > 1
                    || properties.Elements<A.LatinFont>().Count() > 1
                    || properties.Elements<A.EastAsianFont>().Count() > 1
                    || properties.Elements<A.SymbolFont>().Count() > 1)) {
                reason = "A text run contains duplicate color or typeface elements.";
                return false;
            }
            uint masks = 0;
            ushort style = 0;
            bool hasStyle = false;
            AddStyleFlag(properties?.Bold, 0, ref masks, ref style,
                ref hasStyle);
            AddStyleFlag(properties?.Italic, 1, ref masks, ref style,
                ref hasStyle);
            if (properties?.Underline?.HasValue == true) {
                A.TextUnderlineValues value = properties.Underline.Value;
                if (value != A.TextUnderlineValues.None
                    && value != A.TextUnderlineValues.Single) {
                    reason = "Only none and single underline map exactly to the base binary character style.";
                    return false;
                }
                masks |= 1U << 2;
                hasStyle = true;
                if (value == A.TextUnderlineValues.Single) style |= 1 << 2;
            }
            AddStyleFlag(properties?.Kumimoji, 7, ref masks, ref style,
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
            A.LatinFont? latin = properties?.GetFirstChild<A.LatinFont>();
            if (latin != null) {
                if (!HasOnlyAttributes(latin, "typeface")
                    || latin.ChildElements.Count != 0
                    || !fonts.TryGetOrAdd(latin.Typeface?.Value,
                        out ushort index, out reason)) return false;
                latinFont = index;
                masks |= 1U << 16;
            }
            ushort? eastAsianFont = null;
            A.EastAsianFont? eastAsian = properties?
                .GetFirstChild<A.EastAsianFont>();
            if (eastAsian != null) {
                if (!HasOnlyAttributes(eastAsian, "typeface")
                    || eastAsian.ChildElements.Count != 0
                    || !fonts.TryGetOrAdd(eastAsian.Typeface?.Value,
                        out ushort index, out reason)) return false;
                eastAsianFont = index;
                masks |= 1U << 21;
            }
            ushort? symbolFont = null;
            A.SymbolFont? symbol = properties?.GetFirstChild<A.SymbolFont>();
            if (symbol != null) {
                if (!HasOnlyAttributes(symbol, "typeface")
                    || symbol.ChildElements.Count != 0
                    || !fonts.TryGetOrAdd(symbol.Typeface?.Value,
                        out ushort index, out reason)) return false;
                symbolFont = index;
                masks |= 1U << 23;
            }
            short? fontSize = null;
            if (properties?.FontSize?.HasValue == true) {
                int value = properties.FontSize.Value;
                if (value % 100 != 0 || value < 100 || value > 400000) {
                    reason = "A master font size cannot be represented exactly as whole binary PowerPoint points.";
                    return false;
                }
                fontSize = checked((short)(value / 100));
                masks |= 1U << 17;
            }
            byte[]? color = null;
            A.SolidFill? fill = properties?.GetFirstChild<A.SolidFill>();
            if (fill != null) {
                if (!TryReadBinaryColor(fill, out color, out reason)) {
                    return false;
                }
                masks |= 1U << 18;
            }
            short? baseline = null;
            if (properties?.Baseline?.HasValue == true) {
                int value = properties.Baseline.Value;
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

        private static void AddStyleFlag(BooleanValue? value, int bit,
            ref uint masks, ref ushort style, ref bool hasStyle) {
            if (value?.HasValue != true) return;
            masks |= 1U << bit;
            hasStyle = true;
            if (value.Value) style |= checked((ushort)(1 << bit));
        }

        private static bool HasAtMostOneOfEachParagraphChild(
            A.TextParagraphPropertiesType properties) {
            if (properties.Elements<A.DefaultRunProperties>().Count() > 1) {
                return false;
            }
            var seen = new HashSet<Type>();
            foreach (OpenXmlElement child in properties.ChildElements) {
                if (child is A.DefaultRunProperties) continue;
                if (child is not A.LineSpacing and not A.SpaceBefore
                    and not A.SpaceAfter and not A.BulletColorText
                    and not A.BulletColor and not A.BulletSizeText
                    and not A.BulletSizePercentage and not A.BulletSizePoints
                    and not A.BulletFontText and not A.BulletFont
                    and not A.NoBullet and not A.CharacterBullet
                    and not A.AutoNumberedBullet
                    and not A.PictureBullet
                    and not A.TabStopList) return false;
                if (!seen.Add(child.GetType())) return false;
            }
            return CountPresent(properties.GetFirstChild<A.NoBullet>(),
                    properties.GetFirstChild<A.CharacterBullet>(),
                    properties.GetFirstChild<A.AutoNumberedBullet>(),
                    properties.GetFirstChild<A.PictureBullet>()) <= 1
                && CountPresent(properties.GetFirstChild<A.BulletFont>(),
                    properties.GetFirstChild<A.BulletFontText>()) <= 1
                && CountPresent(properties.GetFirstChild<A.BulletColor>(),
                    properties.GetFirstChild<A.BulletColorText>()) <= 1
                && CountPresent(properties.GetFirstChild<A.BulletSizePercentage>(),
                    properties.GetFirstChild<A.BulletSizePoints>(),
                    properties.GetFirstChild<A.BulletSizeText>()) <= 1;
        }

        private static int CountPresent(params OpenXmlElement?[] values) =>
            values.Count(value => value != null);

        private static bool HasOnlyAttributes(OpenXmlElement element,
            params string[] allowed) {
            var names = new HashSet<string>(allowed, StringComparer.Ordinal);
            return element.GetAttributes().All(attribute =>
                names.Contains(attribute.LocalName));
        }

        private static bool TryReadSpacing(OpenXmlCompositeElement? spacing,
            out short? value, out string? reason) {
            value = null;
            reason = null;
            if (spacing == null) return true;
            if (spacing.HasAttributes || spacing.ChildElements.Count != 1) {
                reason = "A master paragraph spacing element contains unsupported content.";
                return false;
            }
            if (spacing.GetFirstChild<A.SpacingPercent>() is { } percent
                && percent.Val?.HasValue == true) {
                if (!HasOnlyAttributes(percent, "val")
                    || percent.ChildElements.Count != 0) {
                    reason = "A master paragraph percentage spacing contains unsupported metadata.";
                    return false;
                }
                int raw = percent.Val.Value;
                if (raw % 1000 != 0 || raw < 0 || raw > 13200000) {
                    reason = "A master paragraph percentage spacing cannot be represented exactly by the binary field.";
                    return false;
                }
                value = checked((short)(raw / 1000));
                return true;
            }
            if (spacing.GetFirstChild<A.SpacingPoints>() is { } points
                && points.Val?.HasValue == true) {
                if (!HasOnlyAttributes(points, "val")
                    || points.ChildElements.Count != 0) {
                    reason = "A master paragraph point spacing contains unsupported metadata.";
                    return false;
                }
                int raw = points.Val.Value;
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

        private static bool TryReadBinaryColor(OpenXmlCompositeElement parent,
            out byte[]? color, out string? reason) {
            color = null;
            reason = null;
            if (parent.HasAttributes || parent.ChildElements.Count != 1) {
                reason = "A master text color contains transforms or unsupported color metadata.";
                return false;
            }
            if (parent.GetFirstChild<A.RgbColorModelHex>() is { } rgb
                && rgb.Val?.Value is string value) {
                string normalized = value.Trim().TrimStart('#');
                if (normalized.Length != 6 || !normalized.All(Uri.IsHexDigit)
                    || (rgb.HasAttributes && !HasOnlyAttributes(rgb, "val"))
                    || rgb.ChildElements.Count != 0) {
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
            if (parent.GetFirstChild<A.SchemeColor>() is { } scheme
                && scheme.Val?.HasValue == true
                && HasOnlyAttributes(scheme, "val")
                && scheme.ChildElements.Count == 0
                && TryMapSchemeColor(scheme.Val.Value, out byte index)) {
                color = new byte[] { 0, 0, 0, index };
                return true;
            }
            reason = "A master text color has no base binary RGB or eight-slot scheme equivalent.";
            return false;
        }

        private static bool TryMapSchemeColor(A.SchemeColorValues value,
            out byte index) {
            if (value == A.SchemeColorValues.Background1
                || value == A.SchemeColorValues.Light1) {
                index = 0;
            } else if (value == A.SchemeColorValues.Text1
                || value == A.SchemeColorValues.Dark1) {
                index = 1;
            } else if (value == A.SchemeColorValues.Accent4) {
                index = 2;
            } else if (value == A.SchemeColorValues.Text2
                || value == A.SchemeColorValues.Dark2) {
                index = 3;
            } else if (value == A.SchemeColorValues.Background2
                || value == A.SchemeColorValues.Light2) {
                index = 4;
            } else if (value == A.SchemeColorValues.Accent1) {
                index = 5;
            } else if (value == A.SchemeColorValues.Accent2) {
                index = 6;
            } else if (value == A.SchemeColorValues.Accent3) {
                index = 7;
            } else {
                index = byte.MaxValue;
                return false;
            }
            return true;
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

        private static bool TryMapAlignment(A.TextAlignmentTypeValues value,
            out ushort mapped) {
            if (value == A.TextAlignmentTypeValues.Left) mapped = 0;
            else if (value == A.TextAlignmentTypeValues.Center) mapped = 1;
            else if (value == A.TextAlignmentTypeValues.Right) mapped = 2;
            else if (value == A.TextAlignmentTypeValues.Justified) mapped = 3;
            else if (value == A.TextAlignmentTypeValues.Distributed) mapped = 4;
            else if (value == A.TextAlignmentTypeValues.ThaiDistributed) mapped = 5;
            else if (value == A.TextAlignmentTypeValues.JustifiedLow) mapped = 6;
            else {
                mapped = ushort.MaxValue;
                return false;
            }
            return true;
        }

        private static bool TryMapFontAlignment(A.TextFontAlignmentValues value,
            out ushort mapped) {
            if (value == A.TextFontAlignmentValues.Baseline) mapped = 0;
            else if (value == A.TextFontAlignmentValues.Top) mapped = 1;
            else if (value == A.TextFontAlignmentValues.Center) mapped = 2;
            else if (value == A.TextFontAlignmentValues.Bottom) mapped = 3;
            else {
                mapped = ushort.MaxValue;
                return false;
            }
            return true;
        }

        private static bool TryMapTabAlignment(A.TextTabAlignmentValues value,
            out ushort mapped) {
            if (value == A.TextTabAlignmentValues.Left) mapped = 0;
            else if (value == A.TextTabAlignmentValues.Center) mapped = 1;
            else if (value == A.TextTabAlignmentValues.Right) mapped = 2;
            else if (value == A.TextTabAlignmentValues.Decimal) mapped = 3;
            else {
                mapped = ushort.MaxValue;
                return false;
            }
            return true;
        }
    }
}
