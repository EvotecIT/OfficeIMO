using System.Globalization;

namespace OfficeIMO.Word.Rtf;

public static partial class WordRtfConverterExtensions {
    private static void CopyPageSetup(WordDocument document, RtfDocument rtf) {
        int? width = ToInt32(checked((uint?)(document.PageSetup.Size.Width?.ToTwips())));
        int? height = ToInt32(checked((uint?)(document.PageSetup.Size.Height?.ToTwips())));
        if (width.HasValue && height.HasValue) {
            rtf.PageSetup.SetPaperSize(width.Value, height.Value);
        } else {
            rtf.PageSetup.PaperWidthTwips = width;
            rtf.PageSetup.PaperHeightTwips = height;
        }

        rtf.PageSetup.SetMargins(
            ToInt32(checked((uint)(document.PageSetup.Margins.Left.GetValueOrDefault().ToTwips()))),
            ToInt32(checked((uint)(document.PageSetup.Margins.Right.GetValueOrDefault().ToTwips()))),
            checked((int?)(document.PageSetup.Margins.Top?.ToTwips())),
            checked((int?)(document.PageSetup.Margins.Bottom?.ToTwips())));
        rtf.PageSetup.SetGutter(ToInt32(checked((uint)(document.PageSetup.Margins.Gutter.GetValueOrDefault().ToTwips()))), document.RtlGutter);
        rtf.PageSetup.SetHeaderFooterDistance(
            ToInt32(checked((uint)(document.PageSetup.Margins.HeaderDistance.GetValueOrDefault().ToTwips()))),
            ToInt32(checked((uint)(document.PageSetup.Margins.FooterDistance.GetValueOrDefault().ToTwips()))));

        rtf.PageSetup.SetLandscape(document.PageSetup.Orientation == OfficePageOrientation.Landscape);
        rtf.PageSetup.SetDifferentFirstPageHeaderFooter(document.PageSetup.DifferentFirstPage);
        OfficeXmlNode sectionProperties = document.Sections[0].Node();
        CopyPageNumbering(sectionProperties.ModelChild("pgNumType"), rtf.PageSetup);
        CopyPageBorders(sectionProperties.ModelChild("pgBorders"), rtf.PageSetup.PageBorders, rtf);
        CopyNoteSettings(
            sectionProperties.ModelChild("footnotePr"),
            sectionProperties.ModelChild("endnotePr"),
            rtf.NoteSettings);
    }

    private static void ApplyPageSetup(RtfDocument rtfDocument, WordDocument document) {
        if (rtfDocument.PageSetup.Landscape) {
            document.PageSetup.Orientation = OfficePageOrientation.Landscape;
        }

        if (rtfDocument.PageSetup.PaperWidthTwips.HasValue) {
            document.PageSetup.Size.Width= ((ToUInt32Value(rtfDocument.PageSetup.PaperWidthTwips.Value)) is { } wordLength2068 ? global::OfficeIMO.OfficeLength.Points((wordLength2068) / 20D) : (global::OfficeIMO.OfficeLength?)null);
        }

        if (rtfDocument.PageSetup.PaperHeightTwips.HasValue) {
            document.PageSetup.Size.Height= ((ToUInt32Value(rtfDocument.PageSetup.PaperHeightTwips.Value)) is { } wordLength2245 ? global::OfficeIMO.OfficeLength.Points((wordLength2245) / 20D) : (global::OfficeIMO.OfficeLength?)null);
        }

        if (rtfDocument.PageSetup.MarginLeftTwips.HasValue) {
            document.PageSetup.Margins.Left = global::OfficeIMO.OfficeLength.Points((ToUInt32Value(rtfDocument.PageSetup.MarginLeftTwips.Value)) / 20D);
        }

        if (rtfDocument.PageSetup.MarginRightTwips.HasValue) {
            document.PageSetup.Margins.Right = global::OfficeIMO.OfficeLength.Points((ToUInt32Value(rtfDocument.PageSetup.MarginRightTwips.Value)) / 20D);
        }

        if (rtfDocument.PageSetup.MarginTopTwips.HasValue) {
            document.PageSetup.Margins.Top= ((rtfDocument.PageSetup.MarginTopTwips.Value) is { } wordLength2754 ? global::OfficeIMO.OfficeLength.Points((wordLength2754) / 20D) : (global::OfficeIMO.OfficeLength?)null);
        }

        if (rtfDocument.PageSetup.MarginBottomTwips.HasValue) {
            document.PageSetup.Margins.Bottom= ((rtfDocument.PageSetup.MarginBottomTwips.Value) is { } wordLength2911 ? global::OfficeIMO.OfficeLength.Points((wordLength2911) / 20D) : (global::OfficeIMO.OfficeLength?)null);
        }

        if (rtfDocument.PageSetup.GutterWidthTwips.HasValue) {
            document.PageSetup.Margins.Gutter = global::OfficeIMO.OfficeLength.Points((ToUInt32Value(rtfDocument.PageSetup.GutterWidthTwips.Value)) / 20D);
        }

        if (rtfDocument.PageSetup.HeaderDistanceTwips.HasValue) {
            document.PageSetup.Margins.HeaderDistance = global::OfficeIMO.OfficeLength.Points((ToUInt32Value(rtfDocument.PageSetup.HeaderDistanceTwips.Value)) / 20D);
        }

        if (rtfDocument.PageSetup.FooterDistanceTwips.HasValue) {
            document.PageSetup.Margins.FooterDistance = global::OfficeIMO.OfficeLength.Points((ToUInt32Value(rtfDocument.PageSetup.FooterDistanceTwips.Value)) / 20D);
        }

        ApplyFirstPageHeaderFooterSelection(document.Sections[0], rtfDocument.PageSetup.DifferentFirstPageHeaderFooter);
        document.RtlGutter = rtfDocument.PageSetup.RtlGutter;
        ApplyPageNumbering(rtfDocument.PageSetup, document.Sections[0]);
        ApplyPageBorders(rtfDocument.PageSetup.PageBorders, document.Sections[0], rtfDocument);
        ApplyNoteSettings(RtfNoteProjection.Resolve(rtfDocument).Settings, document.Sections[0]);
    }

    private static int? ToInt32(uint? value) {
        if (value == null || value.Value > int.MaxValue) {
            return null;
        }

        return (int)value.Value;
    }

    private static int? ToInt32(string? value) {
        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed)
            ? parsed
            : null;
    }

    private static int? ParseUInt16(string? value) =>
        WordXml.ParseUInt32(value) is uint parsed && parsed <= ushort.MaxValue ? (int)parsed : null;

    private static uint ToUInt32Value(int value) {
        if (value < 0) throw new ArgumentOutOfRangeException(nameof(value), "Twip value cannot be negative.");
        return (uint)value;
    }

    private static short ToInt16Value(int value, string parameterName) {
        if (value < 0 || value > short.MaxValue) {
            throw new ArgumentOutOfRangeException(parameterName, "Line-numbering value must fit in a signed 16-bit OpenXML value.");
        }

        return (short)value;
    }

    private static void CopyPageNumbering(OfficeXmlNode source, RtfPageSetup destination) {
        if (source.IsNone) {
            return;
        }

        destination.PageNumberStart = WordXml.ParseInt32(source.WAttr("start"));
        if (destination.PageNumberStart.HasValue) {
            destination.PageNumberRestart = true;
        }

        destination.PageNumberFormat = ToRtfPageNumberFormat(ParseNumberFormat(source.WAttr("fmt")));
    }

    private static void ApplyPageNumbering(RtfPageSetup source, WordSection destination) {
        WordNumberFormat? format = ToWordPageNumberFormat(source.PageNumberFormat);
        int? start = source.PageNumberStart;
        if (source.PageNumberRestart == true && !start.HasValue) {
            start = 1;
        }

        if (start.HasValue || format.HasValue) {
            destination.AddPageNumbering(start, format);
        }
    }

    private static WordNumberFormat? ParseNumberFormat(string? token) =>
        token != null && WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordNumberFormat value) ? value : null;

    private static RtfPageNumberFormat? ToRtfPageNumberFormat(WordNumberFormat? format) {
        if (format == WordNumberFormat.UpperRoman) return RtfPageNumberFormat.UpperRoman;
        if (format == WordNumberFormat.LowerRoman) return RtfPageNumberFormat.LowerRoman;
        if (format == WordNumberFormat.UpperLetter) return RtfPageNumberFormat.UpperLetter;
        if (format == WordNumberFormat.LowerLetter) return RtfPageNumberFormat.LowerLetter;
        if (format == WordNumberFormat.Decimal) return RtfPageNumberFormat.Decimal;
        return null;
    }

    private static WordNumberFormat? ToWordPageNumberFormat(RtfPageNumberFormat? format) {
        if (format == RtfPageNumberFormat.UpperRoman) return WordNumberFormat.UpperRoman;
        if (format == RtfPageNumberFormat.LowerRoman) return WordNumberFormat.LowerRoman;
        if (format == RtfPageNumberFormat.UpperLetter) return WordNumberFormat.UpperLetter;
        if (format == RtfPageNumberFormat.LowerLetter) return WordNumberFormat.LowerLetter;
        if (format == RtfPageNumberFormat.Decimal) return WordNumberFormat.Decimal;
        return null;
    }

    private static void CopyPageBorders(OfficeXmlNode source, RtfPageBorders destination, RtfDocument document) {
        if (source.IsNone) {
            return;
        }

        destination.Scope = ToRtfPageBorderScope(source.WAttr("display"));
        destination.DisplayBehindText = ToRtfPageBorderZOrder(source.WAttr("zOrder"));
        destination.OffsetFrom = ToRtfPageBorderOffset(source.WAttr("offsetFrom"));
        CopyPageBorder(source.ModelChild("top"), destination.Top, document);
        CopyPageBorder(source.ModelChild("bottom"), destination.Bottom, document);
        CopyPageBorder(source.ModelChild("left"), destination.Left, document);
        CopyPageBorder(source.ModelChild("right"), destination.Right, document);
    }

    private static void CopyPageBorder(OfficeXmlNode source, RtfPageBorder destination, RtfDocument document) {
        if (source.IsNone) {
            return;
        }

        destination.Style = ToRtfPageBorderStyle(source.WVal());
        destination.Width = WordXml.ParseUInt32(source.WAttr("sz")) is uint size ? ToRtfBorderWidth(size) : null;
        destination.Space = ToInt32(WordXml.ParseUInt32(source.WAttr("space")));
        destination.Shadow = WordXml.ParseOnOff(source.WAttr("shadow")) ?? false;
        destination.Frame = WordXml.ParseOnOff(source.WAttr("frame")) ?? false;
        if (destination.Shadow && destination.Style == RtfPageBorderStyle.Single) {
            destination.Style = RtfPageBorderStyle.Shadow;
        }

        string? colorHex = source.WAttr("color");
        if (!string.IsNullOrWhiteSpace(colorHex) &&
            !string.Equals(colorHex, "auto", StringComparison.OrdinalIgnoreCase) &&
            TryParseHexColor(colorHex!, out byte red, out byte green, out byte blue)) {
            destination.ColorIndex = GetOrAddColor(document, red, green, blue);
        }
    }

    private static void ApplyPageBorders(RtfPageBorders source, WordSection destination, RtfDocument document) {
        if (!source.HasAnyValue) {
            return;
        }

        OfficeXmlNode sectionProperties = destination.Node();
        OfficeXmlNode pageBorders = sectionProperties.CreateElement(WordXml.W, "pgBorders");
        pageBorders.SetAttribute(WordXml.W, "display", ToWordPageBorderScope(source.Scope));
        pageBorders.SetAttribute(WordXml.W, "zOrder", ToWordPageBorderZOrder(source.DisplayBehindText));
        pageBorders.SetAttribute(WordXml.W, "offsetFrom", ToWordPageBorderOffset(source.OffsetFrom));
        // The borders go in schema order: top, left, bottom, right.
        ApplyPageBorder(source.Top, pageBorders, "top", document);
        ApplyPageBorder(source.Left, pageBorders, "left", document);
        ApplyPageBorder(source.Bottom, pageBorders, "bottom", document);
        ApplyPageBorder(source.Right, pageBorders, "right", document);

        sectionProperties.ModelChild("pgBorders").Remove();
        sectionProperties.Append(pageBorders);
    }

    private static void ApplyPageBorder(RtfPageBorder source, OfficeXmlNode pageBorders, string side, RtfDocument document) {
        if (!source.HasAnyValue) {
            return;
        }

        OfficeXmlNode border = pageBorders.CreateElement(WordXml.W, side);
        border.SetAttribute(WordXml.W, "val", ToWordPageBorderStyle(source.Style).ToOpenXmlToken());
        if (source.Width.HasValue && source.Width.Value >= 0) {
            border.SetAttribute(WordXml.W, "sz", ToWordBorderSize(source.Width.Value).ToString(CultureInfo.InvariantCulture));
        }

        if (source.Space.HasValue && source.Space.Value >= 0) {
            border.SetAttribute(WordXml.W, "space", ((uint)source.Space.Value).ToString(CultureInfo.InvariantCulture));
        }

        if (source.ColorIndex.HasValue) {
            string? color = GetColorHex(document, source.ColorIndex.Value);
            if (!string.IsNullOrWhiteSpace(color)) {
                border.SetAttribute(WordXml.W, "color", color);
            }
        }

        if (source.Shadow || source.Style == RtfPageBorderStyle.Shadow) {
            border.SetAttribute(WordXml.W, "shadow", FormatOnOff(true));
        }

        if (source.Frame) {
            border.SetAttribute(WordXml.W, "frame", FormatOnOff(true));
        }

        pageBorders.Append(border);
    }

    // Page border styles are w:val tokens.
    private static RtfPageBorderStyle ToRtfPageBorderStyle(string? value) {
        if (value == "double") return RtfPageBorderStyle.Double;
        if (value == "dotted") return RtfPageBorderStyle.Dotted;
        if (value == "dashed") return RtfPageBorderStyle.Dashed;
        if (value == "nil" || value == "none") return RtfPageBorderStyle.None;
        if (value == "single") return RtfPageBorderStyle.Single;
        return RtfPageBorderStyle.None;
    }

    private static WordBorderStyle ToWordPageBorderStyle(RtfPageBorderStyle value) {
        switch (value) {
            case RtfPageBorderStyle.Double:
                return WordBorderStyle.Double;
            case RtfPageBorderStyle.Dotted:
                return WordBorderStyle.Dotted;
            case RtfPageBorderStyle.Dashed:
                return WordBorderStyle.Dashed;
            case RtfPageBorderStyle.Single:
            case RtfPageBorderStyle.Shadow:
                return WordBorderStyle.Single;
            default:
                return WordBorderStyle.Nil;
        }
    }

    // Scopes are w:display tokens.
    private static RtfPageBorderScope? ToRtfPageBorderScope(string? value) {
        if (value == "firstPage") return RtfPageBorderScope.FirstPageInSection;
        if (value == "notFirstPage") return RtfPageBorderScope.AllExceptFirstPageInSection;
        if (value == "allPages") return RtfPageBorderScope.AllPagesInSection;
        return null;
    }

    private static string? ToWordPageBorderScope(RtfPageBorderScope? value) {
        if (value == RtfPageBorderScope.FirstPageInSection) return "firstPage";
        if (value == RtfPageBorderScope.AllExceptFirstPageInSection) return "notFirstPage";
        if (value == RtfPageBorderScope.AllPagesInSection || value == RtfPageBorderScope.WholeDocument) return "allPages";
        return null;
    }

    // Stacking is a w:zOrder token.
    private static bool? ToRtfPageBorderZOrder(string? value) {
        if (value == "back") return true;
        if (value == "front") return false;
        return null;
    }

    private static string? ToWordPageBorderZOrder(bool? displayBehindText) {
        if (!displayBehindText.HasValue) return null;
        return displayBehindText.Value ? "back" : "front";
    }

    // Offsets are w:offsetFrom tokens.
    private static RtfPageBorderOffset? ToRtfPageBorderOffset(string? value) {
        if (value == "page") return RtfPageBorderOffset.PageEdge;
        if (value == "text") return RtfPageBorderOffset.Text;
        return null;
    }

    private static string? ToWordPageBorderOffset(RtfPageBorderOffset? value) {
        if (value == RtfPageBorderOffset.PageEdge) return "page";
        if (value == RtfPageBorderOffset.Text) return "text";
        return null;
    }

    private static void CopyNoteSettings(OfficeXmlNode footnoteProperties, OfficeXmlNode endnoteProperties, RtfNoteSettings destination) {
        if (footnoteProperties.Exists) {
            destination.FootnoteStartNumber = ParseUInt16(footnoteProperties.ChildVal("numStart"));
            destination.FootnoteRestart = ToRtfNoteRestart(ParseNoteRestart(footnoteProperties.ChildVal("numRestart")));
            destination.FootnoteNumberFormat = ToRtfNoteNumberFormat(ParseNumberFormat(footnoteProperties.ChildVal("numFmt")));
            destination.FootnotePlacement = ToRtfFootnotePlacement(ParseFootnotePosition(footnoteProperties.ChildVal("pos")));
        }

        if (endnoteProperties.Exists) {
            destination.EndnoteStartNumber = ParseUInt16(endnoteProperties.ChildVal("numStart"));
            destination.EndnoteRestart = ToRtfNoteRestart(ParseNoteRestart(endnoteProperties.ChildVal("numRestart")));
            destination.EndnoteNumberFormat = ToRtfNoteNumberFormat(ParseNumberFormat(endnoteProperties.ChildVal("numFmt")));
            destination.EndnotePlacement = ToRtfEndnotePlacement(ParseEndnotePosition(endnoteProperties.ChildVal("pos")));
        }
    }

    private static void ApplyNoteSettings(RtfNoteSettings source, WordSection destination) {
        bool hasFootnoteSettings = source.FootnoteStartNumber.HasValue || source.FootnoteRestart.HasValue || source.FootnoteNumberFormat.HasValue || source.FootnotePlacement.HasValue;
        bool hasEndnoteSettings = source.EndnoteStartNumber.HasValue || source.EndnoteRestart.HasValue || source.EndnoteNumberFormat.HasValue || source.EndnotePlacement.HasValue;
        if (!hasFootnoteSettings && !hasEndnoteSettings) {
            return;
        }

        if (hasFootnoteSettings) {
            destination.AddFootnoteProperties(
                ToWordNoteNumberFormat(source.FootnoteNumberFormat),
                ToWordFootnotePlacement(source.FootnotePlacement),
                restartNumbering: ToWordNoteRestart(source.FootnoteRestart),
                startNumber: source.FootnoteStartNumber);
        }

        if (hasEndnoteSettings) {
            destination.AddEndnoteProperties(
                ToWordNoteNumberFormat(source.EndnoteNumberFormat),
                ToWordEndnotePlacement(source.EndnotePlacement),
                restartNumbering: source.EndnoteRestart == RtfNoteNumberRestart.EachPage ? null : ToWordNoteRestart(source.EndnoteRestart),
                startNumber: source.EndnoteStartNumber);
        }
    }

    private static WordNoteNumberRestart? ParseNoteRestart(string? token) =>
        token != null && WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordNoteNumberRestart value) ? value : null;

    private static WordFootnotePosition? ParseFootnotePosition(string? token) =>
        token != null && WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordFootnotePosition value) ? value : null;

    private static WordEndnotePosition? ParseEndnotePosition(string? token) =>
        token != null && WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordEndnotePosition value) ? value : null;

    private static RtfNoteNumberFormat? ToRtfNoteNumberFormat(WordNumberFormat? format) {
        if (format == WordNumberFormat.LowerRoman) return RtfNoteNumberFormat.LowerRoman;
        if (format == WordNumberFormat.UpperRoman) return RtfNoteNumberFormat.UpperRoman;
        if (format == WordNumberFormat.LowerLetter) return RtfNoteNumberFormat.LowerLetter;
        if (format == WordNumberFormat.UpperLetter) return RtfNoteNumberFormat.UpperLetter;
        if (format == WordNumberFormat.Decimal) return RtfNoteNumberFormat.Arabic;
        return null;
    }

    private static WordNumberFormat? ToWordNoteNumberFormat(RtfNoteNumberFormat? format) {
        if (format == RtfNoteNumberFormat.LowerRoman) return WordNumberFormat.LowerRoman;
        if (format == RtfNoteNumberFormat.UpperRoman) return WordNumberFormat.UpperRoman;
        if (format == RtfNoteNumberFormat.LowerLetter) return WordNumberFormat.LowerLetter;
        if (format == RtfNoteNumberFormat.UpperLetter) return WordNumberFormat.UpperLetter;
        if (format == RtfNoteNumberFormat.Arabic) return WordNumberFormat.Decimal;
        return null;
    }

    private static RtfNoteNumberRestart? ToRtfNoteRestart(WordNoteNumberRestart? restart) {
        if (restart == WordNoteNumberRestart.EachPage) return RtfNoteNumberRestart.EachPage;
        if (restart == WordNoteNumberRestart.EachSection) return RtfNoteNumberRestart.EachSection;
        if (restart == WordNoteNumberRestart.Continuous) return RtfNoteNumberRestart.Continuous;
        return null;
    }

    private static WordNoteNumberRestart? ToWordNoteRestart(RtfNoteNumberRestart? restart) {
        if (restart == RtfNoteNumberRestart.EachPage) return WordNoteNumberRestart.EachPage;
        if (restart == RtfNoteNumberRestart.EachSection) return WordNoteNumberRestart.EachSection;
        if (restart == RtfNoteNumberRestart.Continuous) return WordNoteNumberRestart.Continuous;
        return null;
    }

    private static void CopyLineNumbering(OfficeXmlNode source, RtfLineNumbering destination) {
        if (source.IsNone) {
            return;
        }

        destination.CountBy = WordXml.ParseInt16(source.WAttr("countBy"));
        destination.StartNumber = WordXml.ParseInt16(source.WAttr("start"));
        destination.DistanceFromTextTwips = ToInt32(source.WAttr("distance"));
        destination.Restart = ToRtfLineNumberRestart(source.WAttr("restart"));
    }

    private static void ApplyLineNumbering(RtfLineNumbering source, WordSection destination) {
        if (!source.HasAnyValue) {
            return;
        }

        OfficeXmlNode sectionProperties = destination.Node();
        OfficeXmlNode lineNumberType = sectionProperties.CreateElement(WordXml.W, "lnNumType");
        lineNumberType.SetAttribute(WordXml.W, "restart", ToWordLineNumberRestart(source.Restart));
        if (source.CountBy.HasValue) {
            lineNumberType.SetAttribute(WordXml.W, "countBy",
                ToInt16Value(source.CountBy.Value, nameof(source.CountBy)).ToString(CultureInfo.InvariantCulture));
        }

        if (source.StartNumber.HasValue) {
            lineNumberType.SetAttribute(WordXml.W, "start",
                ToInt16Value(source.StartNumber.Value, nameof(source.StartNumber)).ToString(CultureInfo.InvariantCulture));
        }

        if (source.DistanceFromTextTwips.HasValue) {
            lineNumberType.SetAttribute(WordXml.W, "distance", source.DistanceFromTextTwips.Value.ToString(CultureInfo.InvariantCulture));
        }

        sectionProperties.RemoveChildren(WordXml.W, "lnNumType");
        sectionProperties.Append(lineNumberType);
    }

    // Restarts are w:restart tokens.
    private static RtfLineNumberRestart? ToRtfLineNumberRestart(string? restart) {
        if (restart == "newPage") return RtfLineNumberRestart.EachPage;
        if (restart == "newSection") return RtfLineNumberRestart.EachSection;
        if (restart == "continuous") return RtfLineNumberRestart.Continuous;
        return null;
    }

    private static string? ToWordLineNumberRestart(RtfLineNumberRestart? restart) {
        if (restart == RtfLineNumberRestart.EachPage) return "newPage";
        if (restart == RtfLineNumberRestart.EachSection) return "newSection";
        if (restart == RtfLineNumberRestart.Continuous) return "continuous";
        return null;
    }

    // Vertical alignments are w:vAlign tokens.
    private static RtfSectionVerticalAlignment? ToRtfSectionVerticalAlignment(string? alignment) {
        if (alignment == "top") return RtfSectionVerticalAlignment.Top;
        if (alignment == "center") return RtfSectionVerticalAlignment.Center;
        if (alignment == "bottom") return RtfSectionVerticalAlignment.Bottom;
        if (alignment == "both") return RtfSectionVerticalAlignment.Justified;
        return null;
    }

    private static void ApplySectionVerticalAlignment(RtfSectionVerticalAlignment? alignment, WordSection destination) {
        string? wordAlignment = ToWordSectionVerticalAlignment(alignment);
        if (wordAlignment == null) {
            return;
        }

        OfficeXmlNode sectionProperties = destination.Node();
        sectionProperties.RemoveChildren(WordXml.W, "vAlign");
        OfficeXmlNode verticalAlignment = sectionProperties.CreateElement(WordXml.W, "vAlign");
        verticalAlignment.SetAttribute(WordXml.W, "val", wordAlignment);
        sectionProperties.Append(verticalAlignment);
    }

    private static string? ToWordSectionVerticalAlignment(RtfSectionVerticalAlignment? alignment) {
        if (alignment == RtfSectionVerticalAlignment.Top) return "top";
        if (alignment == RtfSectionVerticalAlignment.Center) return "center";
        if (alignment == RtfSectionVerticalAlignment.Bottom) return "bottom";
        if (alignment == RtfSectionVerticalAlignment.Justified) return "both";
        return null;
    }

    private static RtfFootnotePlacement? ToRtfFootnotePlacement(WordFootnotePosition? placement) {
        if (placement == WordFootnotePosition.PageBottom) return RtfFootnotePlacement.PageBottom;
        if (placement == WordFootnotePosition.BeneathText) return RtfFootnotePlacement.BeneathText;
        if (placement == WordFootnotePosition.SectionEnd) return RtfFootnotePlacement.SectionEnd;
        return null;
    }

    private static WordFootnotePosition? ToWordFootnotePlacement(RtfFootnotePlacement? placement) {
        if (placement == RtfFootnotePlacement.PageBottom) return WordFootnotePosition.PageBottom;
        if (placement == RtfFootnotePlacement.BeneathText) return WordFootnotePosition.BeneathText;
        if (placement == RtfFootnotePlacement.SectionEnd) return WordFootnotePosition.SectionEnd;
        return null;
    }

    private static RtfEndnotePlacement? ToRtfEndnotePlacement(WordEndnotePosition? placement) {
        if (placement == WordEndnotePosition.SectionEnd) return RtfEndnotePlacement.SectionEnd;
        if (placement == WordEndnotePosition.DocumentEnd) return RtfEndnotePlacement.DocumentEnd;
        return null;
    }

    private static WordEndnotePosition? ToWordEndnotePlacement(RtfEndnotePlacement? placement) {
        if (placement == RtfEndnotePlacement.SectionEnd) return WordEndnotePosition.SectionEnd;
        if (placement == RtfEndnotePlacement.DocumentEnd) return WordEndnotePosition.DocumentEnd;
        return null;
    }

}
