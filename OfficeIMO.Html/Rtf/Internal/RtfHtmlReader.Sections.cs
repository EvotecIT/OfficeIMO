using System.Globalization;

namespace OfficeIMO.Html;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private bool TryStartSection(IElement token) {
            if (!string.Equals(token.LocalName, "section", StringComparison.OrdinalIgnoreCase) ||
                !IsTrue(GetAttribute(token, "data-officeimo-rtf-section"))) {
                return false;
            }

            EndParagraph();
            _currentSection = _document.AddSection();
            _currentSection.ResetLayout();
            ApplySectionLayout(_currentSection, RtfHtmlMetadataCodec.Decode(GetAttribute(token, "data-officeimo-rtf-section-layout")));
            _sectionElementDepth = 1;
            return true;
        }

        private bool TryEndSection(string name) {
            if (_currentSection == null || _sectionElementDepth != 1 || !string.Equals(name, "section", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            EndParagraph();
            _currentSection = null;
            _sectionElementDepth = 0;
            return true;
        }

        private void EnterSectionElement() {
            if (_currentSection != null) {
                _sectionElementDepth++;
            }
        }

        private void ExitSectionElement() {
            if (_currentSection != null && _sectionElementDepth > 1) {
                _sectionElementDepth--;
            }
        }

        private void AddSectionBlock(IRtfBlock block) {
            if (_currentSection != null && _cell == null) {
                _currentSection.AddParsedBlock(block);
            }
        }

        private void ApplyDocumentLayout(Dictionary<string, string> values) {
            ApplyPageSetup(_document.PageSetup, values, "page");
            ApplyNoteSettings(_document.NoteSettings, values, "note");
        }

        private void ApplyColorTable(Dictionary<string, string> values) {
            var colors = new List<RtfColor>();
            for (int index = 0; ; index++) {
                string prefix = "color." + index.ToString(CultureInfo.InvariantCulture);
                int? red = ReadInt(values, prefix + ".red");
                int? green = ReadInt(values, prefix + ".green");
                int? blue = ReadInt(values, prefix + ".blue");
                if (!red.HasValue || !green.HasValue || !blue.HasValue) {
                    break;
                }

                var color = new RtfColor(ToByte(red.Value), ToByte(green.Value), ToByte(blue.Value)) {
                    ThemeColor = ReadEnum<RtfThemeColor>(values, prefix + ".theme"),
                    Tint = ReadInt(values, prefix + ".tint"),
                    Shade = ReadInt(values, prefix + ".shade")
                };
                colors.Add(color);
            }

            if (colors.Count > 0) {
                _document.ReplaceColors(colors);
            }
        }

        private static void ApplySectionLayout(RtfSection section, Dictionary<string, string> values) {
            section.BreakKind = ReadEnum(values, "break", RtfSectionBreakKind.NextPage);
            section.ColumnCount = ReadInt(values, "column.count");
            section.ColumnSpaceTwips = ReadInt(values, "column.space");
            section.ColumnSeparator = ReadBool(values, "column.separator") == true;
            section.VerticalAlignment = ReadEnum<RtfSectionVerticalAlignment>(values, "verticalAlignment");
            section.Direction = ReadEnum<RtfTextDirection>(values, "direction");
            ApplyPageSetup(section.PageSetup, values, "page");
            ApplyNoteSettings(section.NoteSettings, values, "note");
            ApplyLineNumbering(section.LineNumbering, values, "line");

            for (int index = 0; ; index++) {
                string prefix = "column." + index.ToString(CultureInfo.InvariantCulture);
                int? width = ReadInt(values, prefix + ".width");
                int? spaceAfter = ReadInt(values, prefix + ".spaceAfter");
                if (!width.HasValue && !spaceAfter.HasValue) {
                    break;
                }

                section.AddColumn(width, spaceAfter);
            }
        }

        private static void ApplyPageSetup(RtfPageSetup pageSetup, Dictionary<string, string> values, string prefix) {
            pageSetup.PaperWidthTwips = ReadInt(values, prefix + ".paperWidth");
            pageSetup.PaperHeightTwips = ReadInt(values, prefix + ".paperHeight");
            pageSetup.PrinterPaperSize = ReadInt(values, prefix + ".printerPaperSize");
            pageSetup.FirstPagePaperSource = ReadInt(values, prefix + ".firstPagePaperSource");
            pageSetup.OtherPagesPaperSource = ReadInt(values, prefix + ".otherPagesPaperSource");
            pageSetup.MarginLeftTwips = ReadInt(values, prefix + ".marginLeft");
            pageSetup.MarginRightTwips = ReadInt(values, prefix + ".marginRight");
            pageSetup.MarginTopTwips = ReadInt(values, prefix + ".marginTop");
            pageSetup.MarginBottomTwips = ReadInt(values, prefix + ".marginBottom");
            pageSetup.GutterWidthTwips = ReadInt(values, prefix + ".gutter");
            pageSetup.HeaderDistanceTwips = ReadInt(values, prefix + ".headerDistance");
            pageSetup.FooterDistanceTwips = ReadInt(values, prefix + ".footerDistance");
            pageSetup.PageNumberStart = ReadInt(values, prefix + ".pageNumberStart");
            pageSetup.PageNumberRestart = ReadBool(values, prefix + ".pageNumberRestart");
            pageSetup.PageNumberPositionXTwips = ReadInt(values, prefix + ".pageNumberX");
            pageSetup.PageNumberPositionYTwips = ReadInt(values, prefix + ".pageNumberY");
            pageSetup.PageNumberFormat = ReadEnum<RtfPageNumberFormat>(values, prefix + ".pageNumberFormat");
            pageSetup.Landscape = ReadBool(values, prefix + ".landscape") == true;
            pageSetup.DifferentFirstPageHeaderFooter = ReadBool(values, prefix + ".differentFirstPage") == true;
            pageSetup.RtlGutter = ReadBool(values, prefix + ".rtlGutter") == true;
            ApplyPageBorders(pageSetup.PageBorders, values, prefix + ".borders");
        }

        private static void ApplyPageBorders(RtfPageBorders borders, Dictionary<string, string> values, string prefix) {
            borders.IncludeHeader = ReadBool(values, prefix + ".includeHeader") == true;
            borders.IncludeFooter = ReadBool(values, prefix + ".includeFooter") == true;
            borders.SnapToPageBorder = ReadBool(values, prefix + ".snap") == true;
            borders.Scope = ReadEnum<RtfPageBorderScope>(values, prefix + ".scope");
            borders.DisplayBehindText = ReadBool(values, prefix + ".behindText");
            borders.OffsetFrom = ReadEnum<RtfPageBorderOffset>(values, prefix + ".offset");
            ApplyPageBorder(borders.Top, values, prefix + ".top");
            ApplyPageBorder(borders.Bottom, values, prefix + ".bottom");
            ApplyPageBorder(borders.Left, values, prefix + ".left");
            ApplyPageBorder(borders.Right, values, prefix + ".right");
        }

        private static void ApplyPageBorder(RtfPageBorder border, Dictionary<string, string> values, string prefix) {
            border.Style = ReadEnum(values, prefix + ".style", RtfPageBorderStyle.None);
            border.Width = ReadInt(values, prefix + ".width");
            border.Space = ReadInt(values, prefix + ".space");
            border.ColorIndex = ReadInt(values, prefix + ".color");
            border.Shadow = ReadBool(values, prefix + ".shadow") == true;
            border.Frame = ReadBool(values, prefix + ".frame") == true;
        }

        private static void ApplyNoteSettings(RtfNoteSettings settings, Dictionary<string, string> values, string prefix) {
            settings.FootnoteStartNumber = ReadInt(values, prefix + ".footnoteStart");
            settings.FootnoteRestart = ReadEnum<RtfNoteNumberRestart>(values, prefix + ".footnoteRestart");
            settings.FootnoteNumberFormat = ReadEnum<RtfNoteNumberFormat>(values, prefix + ".footnoteFormat");
            settings.FootnotePlacement = ReadEnum<RtfFootnotePlacement>(values, prefix + ".footnotePlacement");
            settings.EndnoteStartNumber = ReadInt(values, prefix + ".endnoteStart");
            settings.EndnoteRestart = ReadEnum<RtfNoteNumberRestart>(values, prefix + ".endnoteRestart");
            settings.EndnoteNumberFormat = ReadEnum<RtfNoteNumberFormat>(values, prefix + ".endnoteFormat");
            settings.EndnotePlacement = ReadEnum<RtfEndnotePlacement>(values, prefix + ".endnotePlacement");
        }

        private static void ApplyLineNumbering(RtfLineNumbering lineNumbering, Dictionary<string, string> values, string prefix) {
            lineNumbering.CountBy = ReadInt(values, prefix + ".countBy");
            lineNumbering.DistanceFromTextTwips = ReadInt(values, prefix + ".distance");
            lineNumbering.StartNumber = ReadInt(values, prefix + ".start");
            lineNumbering.Restart = ReadEnum<RtfLineNumberRestart>(values, prefix + ".restart");
        }

        private static int? ReadInt(Dictionary<string, string> values, string key) {
            return values.TryGetValue(key, out string? value) &&
                   int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed)
                ? parsed
                : null;
        }

        private static bool? ReadBool(Dictionary<string, string> values, string key) {
            if (!values.TryGetValue(key, out string? value)) {
                return null;
            }

            if (bool.TryParse(value, out bool parsed)) {
                return parsed;
            }

            return null;
        }

        private static T? ReadEnum<T>(Dictionary<string, string> values, string key) where T : struct {
            return values.TryGetValue(key, out string? value) &&
                   Enum.TryParse(value, ignoreCase: true, out T parsed)
                ? parsed
                : null;
        }

        private static T ReadEnum<T>(Dictionary<string, string> values, string key, T fallback) where T : struct {
            return ReadEnum<T>(values, key) ?? fallback;
        }

        private static byte ToByte(int value) {
            if (value < byte.MinValue) {
                return byte.MinValue;
            }

            if (value > byte.MaxValue) {
                return byte.MaxValue;
            }

            return (byte)value;
        }
    }
}
