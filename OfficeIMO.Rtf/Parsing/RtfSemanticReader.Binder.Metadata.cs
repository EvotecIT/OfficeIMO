using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

internal static partial class RtfSemanticReader {
    private sealed partial class Binder {
        private static string CollectPlainText(RtfGroup group, int ansiCodePage, int unicodeSkipCount = 1) {
            var builder = new StringBuilder();
            var state = new PlainTextState {
                UnicodeSkipCount = unicodeSkipCount
            };
            foreach (RtfNode node in group.Children) {
                CollectPlainText(node, builder, state, ansiCodePage);
            }

            FlushPendingSurrogate(builder, state);
            return builder.ToString();
        }

        private static string CollectDirectPlainText(IEnumerable<RtfNode> nodes, int ansiCodePage, int unicodeSkipCount) {
            var builder = new StringBuilder();
            var state = new PlainTextState {
                UnicodeSkipCount = unicodeSkipCount
            };

            foreach (RtfNode node in nodes) {
                if (node is RtfGroup) {
                    continue;
                }

                CollectPlainText(node, builder, state, ansiCodePage);
            }

            FlushPendingSurrogate(builder, state);
            return builder.ToString();
        }

        private static void CollectPlainText(RtfNode node, StringBuilder builder, PlainTextState state, int ansiCodePage) {
            switch (node) {
                case RtfText text:
                    AppendAnsiText(text.Text, builder, state, ansiCodePage);
                    break;
                case RtfControlSymbol symbol:
                    if (symbol.Symbol == '\'' && symbol.Parameter.HasValue) {
                        AppendAnsiByte(symbol.Parameter.Value, builder, state, ansiCodePage);
                    } else if (symbol.Symbol == '\\' || symbol.Symbol == '{' || symbol.Symbol == '}') {
                        AppendWithSkip(symbol.Symbol.ToString(), builder, state);
                    } else if (symbol.Symbol == '~') {
                        AppendWithSkip("\u00A0", builder, state);
                    } else if (symbol.Symbol == '_') {
                        AppendWithSkip("\u2011", builder, state);
                    } else if (symbol.Symbol == '-') {
                        AppendWithSkip("\u00AD", builder, state);
                    }
                    break;
                case RtfControlWord control when control.Name == "tab":
                    AppendWithSkip("\t", builder, state);
                    break;
                case RtfControlWord control when control.Name == "line" || control.Name == "par":
                    AppendWithSkip(Environment.NewLine, builder, state);
                    break;
                case RtfControlWord control when IsSpecialCharacterControl(control.Name):
                    AppendWithSkip(GetSpecialCharacterText(control.Name), builder, state);
                    break;
                case RtfControlWord control when control.Name == "uc" && control.Parameter.HasValue && control.Parameter.Value >= 0:
                    state.UnicodeSkipCount = control.Parameter.Value;
                    break;
                case RtfControlWord control when control.Name == "u" && control.Parameter.HasValue:
                    AppendUnicodeValue(control.Parameter.Value, builder, state);
                    state.SkipCharacters = state.UnicodeSkipCount;
                    break;
                case RtfGroup group:
                    RtfGroup? unicodeAlternative = group.Destination == "upr" ? FindUnicodeAlternative(group) : null;
                    if (unicodeAlternative != null) {
                        CollectPlainText(unicodeAlternative, builder, state.Clone(), ansiCodePage);
                        break;
                    }

                    PlainTextState childState = state.Clone();
                    foreach (RtfNode child in group.Children) {
                        CollectPlainText(child, builder, childState, ansiCodePage);
                    }
                    break;
            }
        }

        private static void AppendWithSkip(string text, StringBuilder builder, PlainTextState state) {
            if (state.SkipCharacters <= 0) {
                FlushPendingSurrogate(builder, state);
                builder.Append(text);
                return;
            }

            if (state.SkipCharacters >= text.Length) {
                state.SkipCharacters -= text.Length;
                return;
            }

            FlushPendingSurrogate(builder, state);
            builder.Append(text, state.SkipCharacters, text.Length - state.SkipCharacters);
            state.SkipCharacters = 0;
        }

        private static void AppendAnsiText(string text, StringBuilder builder, PlainTextState state, int ansiCodePage) {
            if (string.IsNullOrEmpty(text)) return;
            int start = 0;
            if (state.PendingAnsiLeadByte.HasValue) {
                if (text[0] <= byte.MaxValue) {
                    byte lead = state.PendingAnsiLeadByte.Value;
                    state.PendingAnsiLeadByte = null;
                    AppendWithSkip(RtfAnsiCodePage.DecodeBytes(ansiCodePage, new[] { lead, (byte)text[0] }), builder, state);
                    start = 1;
                } else {
                    state.PendingAnsiLeadByte = null;
                    AppendWithSkip("\uFFFD", builder, state);
                }
            }

            if (start < text.Length) {
                AppendWithSkip(RtfAnsiCodePage.DecodeText(ansiCodePage, text.Substring(start)), builder, state);
            }
        }

        private static void AppendAnsiByte(int value, StringBuilder builder, PlainTextState state, int ansiCodePage) {
            byte current = (byte)(value & 0xFF);
            if (state.PendingAnsiLeadByte.HasValue) {
                byte lead = state.PendingAnsiLeadByte.Value;
                state.PendingAnsiLeadByte = null;
                AppendWithSkip(RtfAnsiCodePage.DecodeBytes(ansiCodePage, new[] { lead, current }), builder, state);
                return;
            }

            if (RtfAnsiCodePage.IsLeadByte(ansiCodePage, current)) {
                state.PendingAnsiLeadByte = current;
                return;
            }

            AppendWithSkip(RtfAnsiCodePage.DecodeByte(ansiCodePage, current), builder, state);
        }

        private static bool IsSpecialCharacterControl(string controlName) {
            switch (controlName) {
                case "emdash":
                case "endash":
                case "emspace":
                case "enspace":
                case "qmspace":
                case "bullet":
                case "lquote":
                case "rquote":
                case "ldblquote":
                case "rdblquote":
                case "ltrmark":
                case "rtlmark":
                case "zwj":
                case "zwnj":
                    return true;
                default:
                    return false;
            }
        }

        private static void AppendUnicodeValue(int value, StringBuilder builder, PlainTextState state) {
            int unsigned = value < 0 ? value + 65536 : value;
            char codeUnit = (char)unsigned;

            if (char.IsHighSurrogate(codeUnit)) {
                FlushPendingSurrogate(builder, state);
                state.PendingHighSurrogate = codeUnit;
                return;
            }

            if (char.IsLowSurrogate(codeUnit)) {
                if (state.PendingHighSurrogate.HasValue) {
                    builder.Append(state.PendingHighSurrogate.Value);
                    builder.Append(codeUnit);
                    state.PendingHighSurrogate = null;
                } else {
                    builder.Append('\uFFFD');
                }

                return;
            }

            FlushPendingSurrogate(builder, state);
            builder.Append(codeUnit);
        }

        private static void FlushPendingSurrogate(StringBuilder builder, PlainTextState state) {
            if (state.PendingAnsiLeadByte.HasValue) {
                builder.Append('\uFFFD');
                state.PendingAnsiLeadByte = null;
            }

            if (!state.PendingHighSurrogate.HasValue) return;
            builder.Append('\uFFFD');
            state.PendingHighSurrogate = null;
        }

        private sealed class PlainTextState {
            public int UnicodeSkipCount { get; set; } = 1;

            public int SkipCharacters { get; set; }

            public char? PendingHighSurrogate { get; set; }

            public byte? PendingAnsiLeadByte { get; set; }

            public PlainTextState Clone() {
                return new PlainTextState {
                    UnicodeSkipCount = UnicodeSkipCount,
                    SkipCharacters = SkipCharacters,
                    PendingHighSurrogate = PendingHighSurrogate,
                    PendingAnsiLeadByte = PendingAnsiLeadByte
                };
            }
        }

        private static void ReadInfo(RtfGroup root, RtfDocumentInfo info, int ansiCodePage, int unicodeSkipCount) {
            RtfGroup? generatorGroup = root.Children.OfType<RtfGroup>().FirstOrDefault(group => group.Destination == "generator");
            if (generatorGroup != null) {
                info.Generator = EmptyToNull(CollectPlainText(generatorGroup, ansiCodePage, unicodeSkipCount).Trim().TrimEnd(';').Trim());
            }

            RtfGroup? infoGroup = root.Children.OfType<RtfGroup>().FirstOrDefault(group => group.Destination == "info");
            if (infoGroup == null) return;

            foreach (RtfNode node in infoGroup.Children) {
                if (node is RtfGroup child) {
                    string value = CollectPlainText(child, ansiCodePage, unicodeSkipCount).Trim();
                    switch (child.Destination) {
                        case "title":
                            info.Title = value;
                            break;
                        case "subject":
                            info.Subject = value;
                            break;
                        case "author":
                            info.Author = value;
                            break;
                        case "manager":
                            info.Manager = value;
                            break;
                        case "company":
                            info.Company = value;
                            break;
                        case "operator":
                            info.Operator = value;
                            break;
                        case "category":
                            info.Category = value;
                            break;
                        case "keywords":
                            info.Keywords = value;
                            break;
                        case "comment":
                        case "doccomm":
                            info.Comments = value;
                            break;
                        case "hlinkbase":
                            info.HyperlinkBase = value;
                            break;
                        case "creatim":
                            info.Created = ReadInfoTimestamp(child);
                            break;
                        case "revtim":
                            info.Revised = ReadInfoTimestamp(child);
                            break;
                        case "printim":
                            info.Printed = ReadInfoTimestamp(child);
                            break;
                        case "buptim":
                            info.BackedUp = ReadInfoTimestamp(child);
                            break;
                    }
                } else if (node is RtfControlWord control) {
                    switch (control.Name) {
                        case "edmins":
                            info.EditingMinutes = control.Parameter;
                            break;
                        case "nofpages":
                            info.NumberOfPages = control.Parameter;
                            break;
                        case "nofwords":
                            info.NumberOfWords = control.Parameter;
                            break;
                        case "nofchars":
                            info.NumberOfCharacters = control.Parameter;
                            break;
                        case "nofcharsws":
                            info.NumberOfCharactersWithSpaces = control.Parameter;
                            break;
                        case "vern":
                            info.InternalVersion = control.Parameter;
                            break;
                    }
                }
            }
        }

        private static DateTime? ReadInfoTimestamp(RtfGroup group) {
            int? year = null;
            int? month = null;
            int? day = null;
            int hour = 0;
            int minute = 0;
            int second = 0;

            foreach (RtfControlWord control in group.Children.OfType<RtfControlWord>()) {
                switch (control.Name) {
                    case "yr":
                        year = control.Parameter;
                        break;
                    case "mo":
                        month = control.Parameter;
                        break;
                    case "dy":
                        day = control.Parameter;
                        break;
                    case "hr":
                        hour = control.Parameter ?? 0;
                        break;
                    case "min":
                        minute = control.Parameter ?? 0;
                        break;
                    case "sec":
                        second = control.Parameter ?? 0;
                        break;
                }
            }

            if (!year.HasValue || !month.HasValue || !day.HasValue) {
                return null;
            }

            try {
                return new DateTime(year.Value, month.Value, day.Value, hour, minute, second, DateTimeKind.Unspecified);
            } catch (ArgumentOutOfRangeException) {
                return null;
            }
        }

        private static void ReadPageSetup(RtfGroup root, RtfPageSetup pageSetup) {
            RtfPageBorder? currentPageBorder = null;
            foreach (RtfNode child in root.Children) {
                if (!(child is RtfControlWord control)) {
                    continue;
                }

                switch (control.Name) {
                    case "sect":
                    case "sectd":
                    case "pard":
                        return;
                    case "paperw":
                        pageSetup.PaperWidthTwips = control.Parameter;
                        break;
                    case "paperh":
                        pageSetup.PaperHeightTwips = control.Parameter;
                        break;
                    case "psz":
                        pageSetup.PrinterPaperSize = control.Parameter;
                        break;
                    case "binfsxn":
                        pageSetup.FirstPagePaperSource = control.Parameter;
                        break;
                    case "binsxn":
                        pageSetup.OtherPagesPaperSource = control.Parameter;
                        break;
                    case "margl":
                        pageSetup.MarginLeftTwips = control.Parameter;
                        break;
                    case "margr":
                        pageSetup.MarginRightTwips = control.Parameter;
                        break;
                    case "margt":
                        pageSetup.MarginTopTwips = control.Parameter;
                        break;
                    case "margb":
                        pageSetup.MarginBottomTwips = control.Parameter;
                        break;
                    case "gutter":
                    case "guttersxn":
                        pageSetup.GutterWidthTwips = control.Parameter;
                        break;
                    case "headery":
                        pageSetup.HeaderDistanceTwips = control.Parameter;
                        break;
                    case "footery":
                        pageSetup.FooterDistanceTwips = control.Parameter;
                        break;
                    case "rtlgutter":
                        pageSetup.RtlGutter = !control.HasParameter || control.Parameter != 0;
                        break;
                    case "pgnstarts":
                        pageSetup.PageNumberStart = control.Parameter;
                        break;
                    case "pgncont":
                        pageSetup.PageNumberRestart = false;
                        break;
                    case "pgnrestart":
                        pageSetup.PageNumberRestart = true;
                        break;
                    case "pgnx":
                        pageSetup.PageNumberPositionXTwips = control.Parameter;
                        break;
                    case "pgny":
                        pageSetup.PageNumberPositionYTwips = control.Parameter;
                        break;
                    case "pgndec":
                        pageSetup.PageNumberFormat = RtfPageNumberFormat.Decimal;
                        break;
                    case "pgnucrm":
                        pageSetup.PageNumberFormat = RtfPageNumberFormat.UpperRoman;
                        break;
                    case "pgnlcrm":
                        pageSetup.PageNumberFormat = RtfPageNumberFormat.LowerRoman;
                        break;
                    case "pgnucltr":
                        pageSetup.PageNumberFormat = RtfPageNumberFormat.UpperLetter;
                        break;
                    case "pgnlcltr":
                        pageSetup.PageNumberFormat = RtfPageNumberFormat.LowerLetter;
                        break;
                    case "pgndecd":
                        pageSetup.PageNumberFormat = RtfPageNumberFormat.DoubleByteDecimal;
                        break;
                    case "landscape":
                        pageSetup.Landscape = !control.HasParameter || control.Parameter != 0;
                        break;
                    case "titlepg":
                        pageSetup.DifferentFirstPageHeaderFooter = !control.HasParameter || control.Parameter != 0;
                        break;
                    default:
                        TryApplyPageBorderControl(control, pageSetup.PageBorders, ref currentPageBorder);
                        break;
                }
            }
        }

        private static bool TryApplyPageBorderControl(RtfControlWord control, RtfPageBorders pageBorders, ref RtfPageBorder? currentPageBorder) {
            switch (control.Name) {
                case "pgbrdrhead":
                    pageBorders.IncludeHeader = !control.HasParameter || control.Parameter != 0;
                    currentPageBorder = null;
                    return true;
                case "pgbrdrfoot":
                    pageBorders.IncludeFooter = !control.HasParameter || control.Parameter != 0;
                    currentPageBorder = null;
                    return true;
                case "pgbrdrsnap":
                    pageBorders.SnapToPageBorder = !control.HasParameter || control.Parameter != 0;
                    currentPageBorder = null;
                    return true;
                case "pgbrdropt":
                    ApplyPageBorderDisplayOptions(pageBorders, control.Parameter.GetValueOrDefault());
                    currentPageBorder = null;
                    return true;
                case "pgbrdrt":
                    currentPageBorder = pageBorders.Top;
                    currentPageBorder.Style = RtfPageBorderStyle.Single;
                    return true;
                case "pgbrdrb":
                    currentPageBorder = pageBorders.Bottom;
                    currentPageBorder.Style = RtfPageBorderStyle.Single;
                    return true;
                case "pgbrdrl":
                    currentPageBorder = pageBorders.Left;
                    currentPageBorder.Style = RtfPageBorderStyle.Single;
                    return true;
                case "pgbrdrr":
                    currentPageBorder = pageBorders.Right;
                    currentPageBorder.Style = RtfPageBorderStyle.Single;
                    return true;
                case "brdrs":
                    return ApplyPageBorderStyle(currentPageBorder, RtfPageBorderStyle.Single);
                case "brdrdb":
                    return ApplyPageBorderStyle(currentPageBorder, RtfPageBorderStyle.Double);
                case "brdrdot":
                    return ApplyPageBorderStyle(currentPageBorder, RtfPageBorderStyle.Dotted);
                case "brdrdash":
                    return ApplyPageBorderStyle(currentPageBorder, RtfPageBorderStyle.Dashed);
                case "brdrsh":
                    if (currentPageBorder == null) return false;
                    currentPageBorder.Style = RtfPageBorderStyle.Shadow;
                    currentPageBorder.Shadow = true;
                    return true;
                case "brdrnone":
                case "brdrnil":
                    return ApplyPageBorderStyle(currentPageBorder, RtfPageBorderStyle.None);
                case "brdrw":
                    if (currentPageBorder == null) return false;
                    currentPageBorder.Width = control.Parameter;
                    return true;
                case "brsp":
                    if (currentPageBorder == null) return false;
                    currentPageBorder.Space = control.Parameter;
                    return true;
                case "brdrcf":
                    if (currentPageBorder == null) return false;
                    currentPageBorder.ColorIndex = control.Parameter;
                    return true;
                case "brdrframe":
                    if (currentPageBorder == null) return false;
                    currentPageBorder.Frame = !control.HasParameter || control.Parameter != 0;
                    return true;
                default:
                    return false;
            }
        }

        private static bool ApplyPageBorderStyle(RtfPageBorder? border, RtfPageBorderStyle style) {
            if (border == null) return false;
            border.Style = style;
            return true;
        }

        private static void ApplyPageBorderDisplayOptions(RtfPageBorders pageBorders, int value) {
            pageBorders.Scope = (value & 7) switch {
                1 => RtfPageBorderScope.FirstPageInSection,
                2 => RtfPageBorderScope.AllExceptFirstPageInSection,
                3 => RtfPageBorderScope.WholeDocument,
                _ => RtfPageBorderScope.AllPagesInSection
            };
            pageBorders.DisplayBehindText = (value & 8) != 0;
            pageBorders.OffsetFrom = (value & 32) != 0 ? RtfPageBorderOffset.PageEdge : RtfPageBorderOffset.Text;
        }

        private static void ReadNoteSettings(RtfGroup root, RtfNoteSettings noteSettings) {
            foreach (RtfNode child in root.Children) {
                if (!(child is RtfControlWord control)) {
                    continue;
                }

                switch (control.Name) {
                    case "sect":
                    case "sectd":
                    case "pard":
                        return;
                    default:
                        TryApplyNoteSettingsControl(control, noteSettings);
                        break;
                }
            }
        }

        private static bool TryApplyNoteSettingsControl(RtfControlWord control, RtfNoteSettings noteSettings) {
            switch (control.Name) {
                case "ftnstart":
                    noteSettings.FootnoteStartNumber = control.Parameter;
                    return true;
                case "aftnstart":
                    noteSettings.EndnoteStartNumber = control.Parameter;
                    return true;
                case "ftnrstcont":
                    noteSettings.FootnoteRestart = RtfNoteNumberRestart.Continuous;
                    return true;
                case "ftnrestart":
                    noteSettings.FootnoteRestart = RtfNoteNumberRestart.EachSection;
                    return true;
                case "ftnrstpg":
                    noteSettings.FootnoteRestart = RtfNoteNumberRestart.EachPage;
                    return true;
                case "aftnrstcont":
                    noteSettings.EndnoteRestart = RtfNoteNumberRestart.Continuous;
                    return true;
                case "aftnrestart":
                    noteSettings.EndnoteRestart = RtfNoteNumberRestart.EachSection;
                    return true;
                case "ftnnar":
                    noteSettings.FootnoteNumberFormat = RtfNoteNumberFormat.Arabic;
                    return true;
                case "ftnnalc":
                    noteSettings.FootnoteNumberFormat = RtfNoteNumberFormat.LowerLetter;
                    return true;
                case "ftnnauc":
                    noteSettings.FootnoteNumberFormat = RtfNoteNumberFormat.UpperLetter;
                    return true;
                case "ftnnrlc":
                    noteSettings.FootnoteNumberFormat = RtfNoteNumberFormat.LowerRoman;
                    return true;
                case "ftnnruc":
                    noteSettings.FootnoteNumberFormat = RtfNoteNumberFormat.UpperRoman;
                    return true;
                case "ftntj":
                    noteSettings.FootnotePlacement = RtfFootnotePlacement.BeneathText;
                    return true;
                case "ftnbj":
                    noteSettings.FootnotePlacement = RtfFootnotePlacement.PageBottom;
                    return true;
                case "endnotes":
                    noteSettings.FootnotePlacement = RtfFootnotePlacement.SectionEnd;
                    return true;
                case "enddoc":
                    noteSettings.FootnotePlacement = RtfFootnotePlacement.DocumentEnd;
                    return true;
                case "aftnnar":
                    noteSettings.EndnoteNumberFormat = RtfNoteNumberFormat.Arabic;
                    return true;
                case "aftnnalc":
                    noteSettings.EndnoteNumberFormat = RtfNoteNumberFormat.LowerLetter;
                    return true;
                case "aftnnauc":
                    noteSettings.EndnoteNumberFormat = RtfNoteNumberFormat.UpperLetter;
                    return true;
                case "aftnnrlc":
                    noteSettings.EndnoteNumberFormat = RtfNoteNumberFormat.LowerRoman;
                    return true;
                case "aftnnruc":
                    noteSettings.EndnoteNumberFormat = RtfNoteNumberFormat.UpperRoman;
                    return true;
                case "aendnotes":
                    noteSettings.EndnotePlacement = RtfEndnotePlacement.SectionEnd;
                    return true;
                case "aenddoc":
                    noteSettings.EndnotePlacement = RtfEndnotePlacement.DocumentEnd;
                    return true;
                case "aftnbj":
                    noteSettings.EndnotePlacement = RtfEndnotePlacement.PageBottom;
                    return true;
                case "aftntj":
                    noteSettings.EndnotePlacement = RtfEndnotePlacement.BeneathText;
                    return true;
                default:
                    return false;
            }
        }

        private static RtfControlWord? FindAnsiCodePageControl(RtfGroup root) {
            foreach (RtfNode child in root.Children) {
                if (child is RtfControlWord control && control.Name == "ansicpg" && control.Parameter.HasValue) {
                    return control;
                }
            }

            return null;
        }

        private static RtfControlWord? FindUnicodeSkipCountControl(RtfGroup root) {
            foreach (RtfNode child in root.Children) {
                if (child is RtfControlWord control && control.Name == "uc" && control.Parameter.HasValue && control.Parameter.Value >= 0) {
                    return control;
                }
            }

            return null;
        }

        private static RtfDocumentCharacterSet? FindDocumentCharacterSet(RtfGroup root) {
            foreach (RtfNode child in root.Children) {
                if (!(child is RtfControlWord control)) {
                    continue;
                }

                switch (control.Name) {
                    case "ansi":
                        return RtfDocumentCharacterSet.Ansi;
                    case "mac":
                        return RtfDocumentCharacterSet.Mac;
                    case "pc":
                        return RtfDocumentCharacterSet.Pc;
                    case "pca":
                        return RtfDocumentCharacterSet.Pca;
                    case "fonttbl":
                    case "colortbl":
                    case "stylesheet":
                    case "listtable":
                    case "pard":
                    case "sect":
                    case "sectd":
                        return null;
                }
            }

            return null;
        }

    }
}
