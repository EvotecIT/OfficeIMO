using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

internal static partial class RtfSemanticReader {
    private sealed partial class Binder {
        private static IReadOnlyList<RtfStyle> ReadStylesheet(RtfGroup root, int ansiCodePage, int unicodeSkipCount) {
            RtfGroup? stylesheet = root.Children.OfType<RtfGroup>().FirstOrDefault(group => group.Destination == "stylesheet");
            if (stylesheet == null) return Array.Empty<RtfStyle>();

            var styles = new List<RtfStyle>();
            foreach (RtfGroup styleGroup in stylesheet.Children.OfType<RtfGroup>()) {
                RtfStyle? style = ReadStyle(styleGroup, ansiCodePage, unicodeSkipCount);
                if (style != null) {
                    styles.Add(style);
                }
            }

            return styles;
        }

        private static RtfStyle? ReadStyle(RtfGroup styleGroup, int ansiCodePage, int unicodeSkipCount) {
            int? id = null;
            RtfStyleKind kind = RtfStyleKind.Paragraph;
            int? basedOn = null;
            int? next = null;
            int? linked = null;
            RtfStyleKeyCode? keyCode = null;
            bool additive = false;
            bool autoUpdate = false;
            bool hidden = false;
            bool locked = false;
            bool personal = false;
            bool compose = false;
            bool reply = false;
            bool semiHidden = false;
            bool unhideWhenUsed = false;
            bool quickFormat = false;
            int? priority = null;
            int? revisionSaveId = null;
            bool? bold = null;
            bool? italic = null;
            RtfUnderlineStyle? underlineStyle = null;
            double? fontSize = null;
            int? fontId = null;
            int? foregroundColorIndex = null;
            int? highlightColorIndex = null;
            RtfTextAlignment? paragraphAlignment = null;
            RtfTextDirection? paragraphDirection = null;
            int? leftIndent = null;
            int? rightIndent = null;
            int? firstLineIndent = null;
            int? spaceBefore = null;
            int? spaceAfter = null;
            bool? spaceBeforeAuto = null;
            bool? spaceAfterAuto = null;
            int? lineSpacing = null;
            bool? lineSpacingMultiple = null;
            int? backgroundColorIndex = null;
            int? shadingForegroundColorIndex = null;
            int? shadingPatternPercent = null;
            RtfShadingPattern shadingPattern = RtfShadingPattern.None;
            bool? pageBreakBefore = null;
            bool? keepWithNext = null;
            bool? keepLinesTogether = null;
            bool? suppressLineNumbers = null;
            bool? autoHyphenation = null;
            bool? contextualSpacing = null;
            bool? adjustRightIndent = null;
            bool? snapToLineGrid = null;
            bool? widowControl = null;
            int? outlineLevel = null;
            var tabState = new CharacterState();
            RtfParagraphBorderSide? currentBorderSide = null;
            var topBorder = new RtfParagraphBorder();
            var leftBorder = new RtfParagraphBorder();
            var bottomBorder = new RtfParagraphBorder();
            var rightBorder = new RtfParagraphBorder();
            var tableRowFormat = new RtfTableRow();
            var pendingTableCell = new PendingTableCellProperties();
            var rowPadding = new RowBoxMeasurements();
            var rowSpacing = new RowBoxMeasurements();
            RtfTableRowBorderSide? currentTableRowBorderSide = null;

            foreach (RtfNode node in styleGroup.Children) {
                if (node is RtfControlWord control) {
                    if (TryApplyParagraphFrameControl(control, tabState)) {
                        continue;
                    }

                    if (TryApplyTabStopControl(control, tabState)) {
                        continue;
                    }

                    if (TryApplyStyleTableControl(
                        control,
                        tableRowFormat,
                        ref pendingTableCell,
                        rowPadding,
                        rowSpacing,
                        ref currentTableRowBorderSide,
                        ref kind)) {
                        continue;
                    }

                    switch (control.Name) {
                        case "s":
                            id = control.Parameter;
                            kind = RtfStyleKind.Paragraph;
                            break;
                        case "cs":
                            id = control.Parameter;
                            kind = RtfStyleKind.Character;
                            break;
                        case "ts":
                            id = control.Parameter;
                            kind = RtfStyleKind.Table;
                            break;
                        case "sbasedon":
                            basedOn = control.Parameter;
                            break;
                        case "snext":
                            next = control.Parameter;
                            break;
                        case "slink":
                            linked = control.Parameter;
                            break;
                        case "additive":
                            additive = true;
                            break;
                        case "sautoupd":
                            autoUpdate = true;
                            break;
                        case "shidden":
                            hidden = true;
                            break;
                        case "slocked":
                            locked = true;
                            break;
                        case "spersonal":
                            personal = true;
                            break;
                        case "scompose":
                            compose = true;
                            break;
                        case "sreply":
                            reply = true;
                            break;
                        case "ssemihidden":
                            semiHidden = true;
                            break;
                        case "sunhideused":
                            unhideWhenUsed = true;
                            break;
                        case "sqformat":
                            quickFormat = true;
                            break;
                        case "spriority":
                            priority = control.Parameter;
                            break;
                        case "styrsid":
                            revisionSaveId = control.Parameter;
                            break;
                        case "b":
                            bold = !control.HasParameter || control.Parameter != 0;
                            break;
                        case "i":
                            italic = !control.HasParameter || control.Parameter != 0;
                            break;
                        case "fs":
                            if (control.Parameter.HasValue) {
                                fontSize = control.Parameter.Value / 2d;
                            }
                            break;
                        case "f":
                            fontId = control.Parameter;
                            break;
                        case "cf":
                            foregroundColorIndex = control.Parameter;
                            break;
                        case "highlight":
                            highlightColorIndex = control.Parameter;
                            break;
                        case "ul":
                        case "ulw":
                        case "uldb":
                        case "uld":
                        case "uldash":
                        case "uldashd":
                        case "uldashdd":
                        case "ulth":
                        case "ulthd":
                        case "ulthdash":
                        case "ulthdashd":
                        case "ulthdashdd":
                        case "ulwave":
                        case "ulhwave":
                        case "uldbwave":
                        case "ulldash":
                        case "ulthldash":
                        case "ulnone":
                            underlineStyle = ReadStyleUnderline(control);
                            break;
                        case "qc":
                            paragraphAlignment = RtfTextAlignment.Center;
                            break;
                        case "qr":
                            paragraphAlignment = RtfTextAlignment.Right;
                            break;
                        case "qj":
                            paragraphAlignment = RtfTextAlignment.Justify;
                            break;
                        case "ql":
                            paragraphAlignment = RtfTextAlignment.Left;
                            break;
                        case "rtlpar":
                            paragraphDirection = RtfTextDirection.RightToLeft;
                            break;
                        case "ltrpar":
                            paragraphDirection = RtfTextDirection.LeftToRight;
                            break;
                        case "li":
                            leftIndent = control.Parameter;
                            break;
                        case "ri":
                            rightIndent = control.Parameter;
                            break;
                        case "fi":
                            firstLineIndent = control.Parameter;
                            break;
                        case "sb":
                            spaceBefore = control.Parameter;
                            break;
                        case "sa":
                            spaceAfter = control.Parameter;
                            break;
                        case "sbauto":
                            spaceBeforeAuto = !control.HasParameter || control.Parameter != 0;
                            break;
                        case "saauto":
                            spaceAfterAuto = !control.HasParameter || control.Parameter != 0;
                            break;
                        case "sl":
                            lineSpacing = control.Parameter;
                            break;
                        case "slmult":
                            lineSpacingMultiple = !control.HasParameter || control.Parameter != 0;
                            break;
                        case "cbpat":
                            backgroundColorIndex = control.Parameter;
                            break;
                        case "cfpat":
                            shadingForegroundColorIndex = control.Parameter;
                            break;
                        case "shading":
                            shadingPatternPercent = control.Parameter;
                            break;
                        case "bghoriz":
                        case "bgvert":
                        case "bgfdiag":
                        case "bgbdiag":
                        case "bgcross":
                        case "bgdcross":
                        case "bgdkhoriz":
                        case "bgdkvert":
                        case "bgdkfdiag":
                        case "bgdkbdiag":
                        case "bgdkcross":
                        case "bgdkdcross":
                            shadingPattern = ReadShadingPattern(control.Name);
                            break;
                        case "pagebb":
                            pageBreakBefore = !control.HasParameter || control.Parameter != 0;
                            break;
                        case "keepn":
                            keepWithNext = !control.HasParameter || control.Parameter != 0;
                            break;
                        case "keep":
                            keepLinesTogether = !control.HasParameter || control.Parameter != 0;
                            break;
                        case "noline":
                            suppressLineNumbers = !control.HasParameter || control.Parameter != 0;
                            break;
                        case "hyphpar":
                            autoHyphenation = !control.HasParameter || control.Parameter != 0;
                            break;
                        case "contextualspace":
                            contextualSpacing = !control.HasParameter || control.Parameter != 0;
                            break;
                        case "adjustright":
                            adjustRightIndent = !control.HasParameter || control.Parameter != 0;
                            break;
                        case "nosnaplinegrid":
                            snapToLineGrid = control.HasParameter && control.Parameter == 0;
                            break;
                        case "widctlpar":
                            widowControl = !control.HasParameter || control.Parameter != 0;
                            break;
                        case "nowidctlpar":
                            widowControl = false;
                            break;
                        case "outlinelevel":
                            outlineLevel = control.Parameter;
                            break;
                        case "brdrt":
                            currentBorderSide = RtfParagraphBorderSide.Top;
                            currentTableRowBorderSide = null;
                            pendingTableCell.CurrentBorderSide = null;
                            GetStyleBorder(topBorder, leftBorder, bottomBorder, rightBorder, currentBorderSide.Value).Style = RtfParagraphBorderStyle.Single;
                            break;
                        case "brdrl":
                            currentBorderSide = RtfParagraphBorderSide.Left;
                            currentTableRowBorderSide = null;
                            pendingTableCell.CurrentBorderSide = null;
                            GetStyleBorder(topBorder, leftBorder, bottomBorder, rightBorder, currentBorderSide.Value).Style = RtfParagraphBorderStyle.Single;
                            break;
                        case "brdrb":
                            currentBorderSide = RtfParagraphBorderSide.Bottom;
                            currentTableRowBorderSide = null;
                            pendingTableCell.CurrentBorderSide = null;
                            GetStyleBorder(topBorder, leftBorder, bottomBorder, rightBorder, currentBorderSide.Value).Style = RtfParagraphBorderStyle.Single;
                            break;
                        case "brdrr":
                            currentBorderSide = RtfParagraphBorderSide.Right;
                            currentTableRowBorderSide = null;
                            pendingTableCell.CurrentBorderSide = null;
                            GetStyleBorder(topBorder, leftBorder, bottomBorder, rightBorder, currentBorderSide.Value).Style = RtfParagraphBorderStyle.Single;
                            break;
                        case "brdrs":
                            ApplyStyleBorderStyle(topBorder, leftBorder, bottomBorder, rightBorder, currentBorderSide, RtfParagraphBorderStyle.Single);
                            break;
                        case "brdrdb":
                            ApplyStyleBorderStyle(topBorder, leftBorder, bottomBorder, rightBorder, currentBorderSide, RtfParagraphBorderStyle.Double);
                            break;
                        case "brdrdot":
                            ApplyStyleBorderStyle(topBorder, leftBorder, bottomBorder, rightBorder, currentBorderSide, RtfParagraphBorderStyle.Dotted);
                            break;
                        case "brdrdash":
                            ApplyStyleBorderStyle(topBorder, leftBorder, bottomBorder, rightBorder, currentBorderSide, RtfParagraphBorderStyle.Dashed);
                            break;
                        case "brdrnil":
                        case "brdrnone":
                            ApplyStyleBorderStyle(topBorder, leftBorder, bottomBorder, rightBorder, currentBorderSide, RtfParagraphBorderStyle.None);
                            break;
                        case "brdrw":
                            ApplyStyleBorderWidth(topBorder, leftBorder, bottomBorder, rightBorder, currentBorderSide, control.Parameter);
                            break;
                        case "brdrcf":
                            ApplyStyleBorderColor(topBorder, leftBorder, bottomBorder, rightBorder, currentBorderSide, control.Parameter);
                            break;
                    }
                } else if (node is RtfGroup childGroup) {
                    if (childGroup.Destination == "keycode") {
                        keyCode = ReadStyleKeyCode(childGroup, ansiCodePage, unicodeSkipCount);
                    } else if (childGroup.Destination == "pn") {
                        tabState.AnsiCodePage = ansiCodePage;
                        tabState.UnicodeSkipCount = unicodeSkipCount;
                        ReadLegacyNumbering(childGroup, tabState);
                    }
                }
            }

            if (!id.HasValue) {
                return null;
            }

            string name = CollectDirectPlainText(styleGroup.Children, ansiCodePage, unicodeSkipCount).Trim().TrimEnd(';').Trim();
            var style = new RtfStyle(id.Value, name, kind) {
                BasedOnStyleId = basedOn,
                NextStyleId = next,
                LinkedStyleId = linked,
                KeyCode = keyCode,
                Additive = additive,
                AutoUpdate = autoUpdate,
                Hidden = hidden,
                Locked = locked,
                Personal = personal,
                Compose = compose,
                Reply = reply,
                SemiHidden = semiHidden,
                UnhideWhenUsed = unhideWhenUsed,
                QuickFormat = quickFormat,
                Priority = priority,
                RevisionSaveId = revisionSaveId,
                Bold = bold,
                Italic = italic,
                UnderlineStyle = underlineStyle,
                FontSize = fontSize,
                FontId = fontId,
                ForegroundColorIndex = foregroundColorIndex,
                HighlightColorIndex = highlightColorIndex,
                ParagraphAlignment = paragraphAlignment,
                ParagraphDirection = paragraphDirection,
                LeftIndentTwips = leftIndent,
                RightIndentTwips = rightIndent,
                FirstLineIndentTwips = firstLineIndent,
                SpaceBeforeTwips = spaceBefore,
                SpaceAfterTwips = spaceAfter,
                SpaceBeforeAuto = spaceBeforeAuto,
                SpaceAfterAuto = spaceAfterAuto,
                LineSpacingTwips = lineSpacing,
                LineSpacingMultiple = lineSpacingMultiple,
                BackgroundColorIndex = backgroundColorIndex,
                ShadingForegroundColorIndex = shadingForegroundColorIndex,
                ShadingPatternPercent = shadingPatternPercent,
                ShadingPattern = shadingPattern,
                PageBreakBefore = pageBreakBefore,
                KeepWithNext = keepWithNext,
                KeepLinesTogether = keepLinesTogether,
                SuppressLineNumbers = suppressLineNumbers,
                AutoHyphenation = autoHyphenation,
                ContextualSpacing = contextualSpacing,
                AdjustRightIndent = adjustRightIndent,
                SnapToLineGrid = snapToLineGrid,
                WidowControl = widowControl,
                OutlineLevel = outlineLevel
            };
            style.ReplaceTabStops(tabState.TabStops);
            style.Frame.CopyFrom(tabState.Frame);
            style.LegacyNumbering.CopyFrom(tabState.LegacyNumbering);
            CopyStyleBorder(topBorder, style.TopBorder);
            CopyStyleBorder(leftBorder, style.LeftBorder);
            CopyStyleBorder(bottomBorder, style.BottomBorder);
            CopyStyleBorder(rightBorder, style.RightBorder);
            CopyStyleTableRowFormat(tableRowFormat, style.TableRowFormat);
            return style;
        }

        private static RtfUnderlineStyle ReadStyleUnderline(RtfControlWord control) {
            if (control.Name != "ulnone" && control.HasParameter && control.Parameter == 0) {
                return RtfUnderlineStyle.None;
            }

            return control.Name switch {
                "ul" => RtfUnderlineStyle.Single,
                "ulw" => RtfUnderlineStyle.Words,
                "uldb" => RtfUnderlineStyle.Double,
                "uld" => RtfUnderlineStyle.Dotted,
                "uldash" => RtfUnderlineStyle.Dash,
                "uldashd" => RtfUnderlineStyle.DashDot,
                "uldashdd" => RtfUnderlineStyle.DashDotDot,
                "ulth" => RtfUnderlineStyle.Thick,
                "ulthd" => RtfUnderlineStyle.ThickDotted,
                "ulthdash" => RtfUnderlineStyle.ThickDash,
                "ulthdashd" => RtfUnderlineStyle.ThickDashDot,
                "ulthdashdd" => RtfUnderlineStyle.ThickDashDotDot,
                "ulwave" => RtfUnderlineStyle.Wave,
                "ulhwave" => RtfUnderlineStyle.HeavyWave,
                "uldbwave" => RtfUnderlineStyle.DoubleWave,
                "ulldash" => RtfUnderlineStyle.LongDash,
                "ulthldash" => RtfUnderlineStyle.ThickLongDash,
                _ => RtfUnderlineStyle.None
            };
        }

        private static RtfShadingPattern ReadShadingPattern(string controlName) {
            return controlName switch {
                "bghoriz" => RtfShadingPattern.Horizontal,
                "bgvert" => RtfShadingPattern.Vertical,
                "bgfdiag" => RtfShadingPattern.ForwardDiagonal,
                "bgbdiag" => RtfShadingPattern.BackwardDiagonal,
                "bgcross" => RtfShadingPattern.Cross,
                "bgdcross" => RtfShadingPattern.DiagonalCross,
                "bgdkhoriz" => RtfShadingPattern.DarkHorizontal,
                "bgdkvert" => RtfShadingPattern.DarkVertical,
                "bgdkfdiag" => RtfShadingPattern.DarkForwardDiagonal,
                "bgdkbdiag" => RtfShadingPattern.DarkBackwardDiagonal,
                "bgdkcross" => RtfShadingPattern.DarkCross,
                "bgdkdcross" => RtfShadingPattern.DarkDiagonalCross,
                _ => RtfShadingPattern.None
            };
        }

        private static RtfParagraphBorder GetStyleBorder(RtfParagraphBorder topBorder, RtfParagraphBorder leftBorder, RtfParagraphBorder bottomBorder, RtfParagraphBorder rightBorder, RtfParagraphBorderSide side) {
            switch (side) {
                case RtfParagraphBorderSide.Top:
                    return topBorder;
                case RtfParagraphBorderSide.Left:
                    return leftBorder;
                case RtfParagraphBorderSide.Bottom:
                    return bottomBorder;
                default:
                    return rightBorder;
            }
        }

        private static void ApplyStyleBorderStyle(RtfParagraphBorder topBorder, RtfParagraphBorder leftBorder, RtfParagraphBorder bottomBorder, RtfParagraphBorder rightBorder, RtfParagraphBorderSide? currentBorderSide, RtfParagraphBorderStyle style) {
            if (!currentBorderSide.HasValue) return;
            GetStyleBorder(topBorder, leftBorder, bottomBorder, rightBorder, currentBorderSide.Value).Style = style;
        }

        private static void ApplyStyleBorderWidth(RtfParagraphBorder topBorder, RtfParagraphBorder leftBorder, RtfParagraphBorder bottomBorder, RtfParagraphBorder rightBorder, RtfParagraphBorderSide? currentBorderSide, int? width) {
            if (!currentBorderSide.HasValue) return;
            GetStyleBorder(topBorder, leftBorder, bottomBorder, rightBorder, currentBorderSide.Value).Width = width;
        }

        private static void ApplyStyleBorderColor(RtfParagraphBorder topBorder, RtfParagraphBorder leftBorder, RtfParagraphBorder bottomBorder, RtfParagraphBorder rightBorder, RtfParagraphBorderSide? currentBorderSide, int? colorIndex) {
            if (!currentBorderSide.HasValue) return;
            GetStyleBorder(topBorder, leftBorder, bottomBorder, rightBorder, currentBorderSide.Value).ColorIndex = colorIndex;
        }

        private static void CopyStyleBorder(RtfParagraphBorder source, RtfParagraphBorder destination) {
            destination.Style = source.Style;
            destination.Width = source.Width;
            destination.ColorIndex = source.ColorIndex;
        }

        private static RtfStyleKeyCode ReadStyleKeyCode(RtfGroup group, int ansiCodePage, int unicodeSkipCount) {
            var keyCode = new RtfStyleKeyCode {
                Key = EmptyToNull(CollectPlainText(group, ansiCodePage, unicodeSkipCount).Trim())
            };

            foreach (RtfControlWord control in group.Children.OfType<RtfControlWord>()) {
                switch (control.Name) {
                    case "shift":
                        keyCode.Shift = true;
                        break;
                    case "ctrl":
                        keyCode.Control = true;
                        break;
                    case "alt":
                        keyCode.Alt = true;
                        break;
                    case "fn":
                        keyCode.FunctionKey = control.Parameter;
                        break;
                }
            }

            return keyCode;
        }
    }
}
