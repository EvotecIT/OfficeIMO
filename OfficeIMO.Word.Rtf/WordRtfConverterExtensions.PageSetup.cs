using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Rtf;

public static partial class WordRtfConverterExtensions {
    private static void CopyPageSetup(WordDocument document, RtfDocument rtf) {
        int? width = ToInt32(document.PageSettings.Width);
        int? height = ToInt32(document.PageSettings.Height);
        if (width.HasValue && height.HasValue) {
            rtf.PageSetup.SetPaperSize(width.Value, height.Value);
        } else {
            rtf.PageSetup.PaperWidthTwips = width;
            rtf.PageSetup.PaperHeightTwips = height;
        }

        rtf.PageSetup.SetMargins(
            ToInt32(document.Margins.Left),
            ToInt32(document.Margins.Right),
            document.Margins.Top,
            document.Margins.Bottom);
        rtf.PageSetup.SetGutter(ToInt32(document.Margins.Gutter), document.RtlGutter);
        rtf.PageSetup.SetHeaderFooterDistance(
            ToInt32(document.Margins.HeaderDistance),
            ToInt32(document.Margins.FooterDistance));

        rtf.PageSetup.SetLandscape(document.PageOrientation == PageOrientationValues.Landscape);
        rtf.PageSetup.SetDifferentFirstPageHeaderFooter(document.DifferentFirstPage);
        CopyPageNumbering(document.Sections[0]._sectionProperties.GetFirstChild<PageNumberType>(), rtf.PageSetup);
        CopyPageBorders(document.Sections[0]._sectionProperties.GetFirstChild<PageBorders>(), rtf.PageSetup.PageBorders, rtf);
        CopyNoteSettings(
            document.Sections[0]._sectionProperties.GetFirstChild<FootnoteProperties>(),
            document.Sections[0]._sectionProperties.GetFirstChild<EndnoteProperties>(),
            rtf.NoteSettings);
    }

    private static void ApplyPageSetup(RtfDocument rtfDocument, WordDocument document) {
        if (rtfDocument.PageSetup.Landscape) {
            document.PageOrientation = PageOrientationValues.Landscape;
        }

        if (rtfDocument.PageSetup.PaperWidthTwips.HasValue) {
            document.PageSettings.Width = ToUInt32Value(rtfDocument.PageSetup.PaperWidthTwips.Value);
        }

        if (rtfDocument.PageSetup.PaperHeightTwips.HasValue) {
            document.PageSettings.Height = ToUInt32Value(rtfDocument.PageSetup.PaperHeightTwips.Value);
        }

        if (rtfDocument.PageSetup.MarginLeftTwips.HasValue) {
            document.Margins.Left = ToUInt32Value(rtfDocument.PageSetup.MarginLeftTwips.Value);
        }

        if (rtfDocument.PageSetup.MarginRightTwips.HasValue) {
            document.Margins.Right = ToUInt32Value(rtfDocument.PageSetup.MarginRightTwips.Value);
        }

        if (rtfDocument.PageSetup.MarginTopTwips.HasValue) {
            document.Margins.Top = rtfDocument.PageSetup.MarginTopTwips.Value;
        }

        if (rtfDocument.PageSetup.MarginBottomTwips.HasValue) {
            document.Margins.Bottom = rtfDocument.PageSetup.MarginBottomTwips.Value;
        }

        if (rtfDocument.PageSetup.GutterWidthTwips.HasValue) {
            document.Margins.Gutter = ToUInt32Value(rtfDocument.PageSetup.GutterWidthTwips.Value);
        }

        if (rtfDocument.PageSetup.HeaderDistanceTwips.HasValue) {
            document.Margins.HeaderDistance = ToUInt32Value(rtfDocument.PageSetup.HeaderDistanceTwips.Value);
        }

        if (rtfDocument.PageSetup.FooterDistanceTwips.HasValue) {
            document.Margins.FooterDistance = ToUInt32Value(rtfDocument.PageSetup.FooterDistanceTwips.Value);
        }

        document.DifferentFirstPage = rtfDocument.PageSetup.DifferentFirstPageHeaderFooter;
        document.RtlGutter = rtfDocument.PageSetup.RtlGutter;
        ApplyPageNumbering(rtfDocument.PageSetup, document.Sections[0]);
        ApplyPageBorders(rtfDocument.PageSetup.PageBorders, document.Sections[0], rtfDocument);
        ApplyNoteSettings(rtfDocument.NoteSettings, document.Sections[0]);
    }

    private static int? ToInt32(UInt32Value? value) {
        if (value?.Value == null || value.Value > int.MaxValue) {
            return null;
        }

        return (int)value.Value;
    }

    private static int? ToInt32(UInt16Value? value) {
        if (value?.Value == null) {
            return null;
        }

        return value.Value;
    }

    private static int? ToInt32(Int16Value? value) {
        if (value?.Value == null) {
            return null;
        }

        return value.Value;
    }

    private static int? ToInt32(StringValue? value) {
        return int.TryParse(value?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed)
            ? parsed
            : null;
    }

    private static UInt32Value ToUInt32Value(int value) {
        if (value < 0) throw new ArgumentOutOfRangeException(nameof(value), "Twip value cannot be negative.");
        return (UInt32Value)(uint)value;
    }

    private static Int16Value ToInt16Value(int value, string parameterName) {
        if (value < 0 || value > short.MaxValue) {
            throw new ArgumentOutOfRangeException(parameterName, "Line-numbering value must fit in a signed 16-bit OpenXML value.");
        }

        return (Int16Value)(short)value;
    }

    private static void CopyPageNumbering(PageNumberType? source, RtfPageSetup destination) {
        if (source == null) {
            return;
        }

        destination.PageNumberStart = source.Start?.Value;
        if (destination.PageNumberStart.HasValue) {
            destination.PageNumberRestart = true;
        }

        destination.PageNumberFormat = ToRtfPageNumberFormat(source.Format?.Value);
    }

    private static void ApplyPageNumbering(RtfPageSetup source, WordSection destination) {
        NumberFormatValues? format = ToWordPageNumberFormat(source.PageNumberFormat);
        int? start = source.PageNumberStart;
        if (source.PageNumberRestart == true && !start.HasValue) {
            start = 1;
        }

        if (start.HasValue || format.HasValue) {
            destination.AddPageNumbering(start, format);
        }
    }

    private static RtfPageNumberFormat? ToRtfPageNumberFormat(NumberFormatValues? format) {
        if (format == NumberFormatValues.UpperRoman) return RtfPageNumberFormat.UpperRoman;
        if (format == NumberFormatValues.LowerRoman) return RtfPageNumberFormat.LowerRoman;
        if (format == NumberFormatValues.UpperLetter) return RtfPageNumberFormat.UpperLetter;
        if (format == NumberFormatValues.LowerLetter) return RtfPageNumberFormat.LowerLetter;
        if (format == NumberFormatValues.Decimal) return RtfPageNumberFormat.Decimal;
        return null;
    }

    private static NumberFormatValues? ToWordPageNumberFormat(RtfPageNumberFormat? format) {
        if (format == RtfPageNumberFormat.UpperRoman) return NumberFormatValues.UpperRoman;
        if (format == RtfPageNumberFormat.LowerRoman) return NumberFormatValues.LowerRoman;
        if (format == RtfPageNumberFormat.UpperLetter) return NumberFormatValues.UpperLetter;
        if (format == RtfPageNumberFormat.LowerLetter) return NumberFormatValues.LowerLetter;
        if (format == RtfPageNumberFormat.Decimal) return NumberFormatValues.Decimal;
        return null;
    }

    private static void CopyPageBorders(PageBorders? source, RtfPageBorders destination, RtfDocument document) {
        if (source == null) {
            return;
        }

        destination.Scope = ToRtfPageBorderScope(source.Display?.Value);
        destination.DisplayBehindText = ToRtfPageBorderZOrder(source.ZOrder?.Value);
        destination.OffsetFrom = ToRtfPageBorderOffset(source.OffsetFrom?.Value);
        CopyPageBorder(source.TopBorder, destination.Top, document);
        CopyPageBorder(source.BottomBorder, destination.Bottom, document);
        CopyPageBorder(source.LeftBorder, destination.Left, document);
        CopyPageBorder(source.RightBorder, destination.Right, document);
    }

    private static void CopyPageBorder(BorderType? source, RtfPageBorder destination, RtfDocument document) {
        if (source == null) {
            return;
        }

        destination.Style = ToRtfPageBorderStyle(source.Val?.Value);
        destination.Width = ToInt32(source.Size);
        destination.Space = ToInt32(source.Space);
        destination.Shadow = source.Shadow?.Value ?? false;
        destination.Frame = source.Frame?.Value ?? false;
        if (destination.Shadow && destination.Style == RtfPageBorderStyle.Single) {
            destination.Style = RtfPageBorderStyle.Shadow;
        }

        string? colorHex = source.Color?.Value;
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

        var pageBorders = new PageBorders {
            Display = ToWordPageBorderScope(source.Scope),
            ZOrder = ToWordPageBorderZOrder(source.DisplayBehindText),
            OffsetFrom = ToWordPageBorderOffset(source.OffsetFrom)
        };
        ApplyPageBorder(source.Top, new TopBorder(), border => pageBorders.TopBorder = border, document);
        ApplyPageBorder(source.Bottom, new BottomBorder(), border => pageBorders.BottomBorder = border, document);
        ApplyPageBorder(source.Left, new LeftBorder(), border => pageBorders.LeftBorder = border, document);
        ApplyPageBorder(source.Right, new RightBorder(), border => pageBorders.RightBorder = border, document);

        PageBorders? existing = destination._sectionProperties.GetFirstChild<PageBorders>();
        existing?.Remove();
        destination._sectionProperties.Append(pageBorders);
    }

    private static void ApplyPageBorder<TBorder>(RtfPageBorder source, TBorder border, Action<TBorder> setBorder, RtfDocument document)
        where TBorder : BorderType {
        if (!source.HasAnyValue) {
            return;
        }

        border.Val = ToWordPageBorderStyle(source.Style);
        if (source.Width.HasValue && source.Width.Value >= 0) {
            border.Size = (UInt32Value)(uint)source.Width.Value;
        }

        if (source.Space.HasValue && source.Space.Value >= 0) {
            border.Space = (UInt32Value)(uint)source.Space.Value;
        }

        if (source.ColorIndex.HasValue) {
            string? color = GetColorHex(document, source.ColorIndex.Value);
            if (!string.IsNullOrWhiteSpace(color)) {
                border.Color = color;
            }
        }

        if (source.Shadow || source.Style == RtfPageBorderStyle.Shadow) {
            border.Shadow = true;
        }

        if (source.Frame) {
            border.Frame = true;
        }

        setBorder(border);
    }

    private static RtfPageBorderStyle ToRtfPageBorderStyle(BorderValues? value) {
        if (value == BorderValues.Double) return RtfPageBorderStyle.Double;
        if (value == BorderValues.Dotted) return RtfPageBorderStyle.Dotted;
        if (value == BorderValues.Dashed) return RtfPageBorderStyle.Dashed;
        if (value == BorderValues.Nil || value == BorderValues.None) return RtfPageBorderStyle.None;
        if (value == BorderValues.Single) return RtfPageBorderStyle.Single;
        return RtfPageBorderStyle.None;
    }

    private static BorderValues? ToWordPageBorderStyle(RtfPageBorderStyle value) {
        switch (value) {
            case RtfPageBorderStyle.Double:
                return BorderValues.Double;
            case RtfPageBorderStyle.Dotted:
                return BorderValues.Dotted;
            case RtfPageBorderStyle.Dashed:
                return BorderValues.Dashed;
            case RtfPageBorderStyle.Single:
            case RtfPageBorderStyle.Shadow:
                return BorderValues.Single;
            default:
                return BorderValues.Nil;
        }
    }

    private static RtfPageBorderScope? ToRtfPageBorderScope(PageBorderDisplayValues? value) {
        if (value == PageBorderDisplayValues.FirstPage) return RtfPageBorderScope.FirstPageInSection;
        if (value == PageBorderDisplayValues.NotFirstPage) return RtfPageBorderScope.AllExceptFirstPageInSection;
        if (value == PageBorderDisplayValues.AllPages) return RtfPageBorderScope.AllPagesInSection;
        return null;
    }

    private static PageBorderDisplayValues? ToWordPageBorderScope(RtfPageBorderScope? value) {
        if (value == RtfPageBorderScope.FirstPageInSection) return PageBorderDisplayValues.FirstPage;
        if (value == RtfPageBorderScope.AllExceptFirstPageInSection) return PageBorderDisplayValues.NotFirstPage;
        if (value == RtfPageBorderScope.AllPagesInSection || value == RtfPageBorderScope.WholeDocument) return PageBorderDisplayValues.AllPages;
        return null;
    }

    private static bool? ToRtfPageBorderZOrder(PageBorderZOrderValues? value) {
        if (value == PageBorderZOrderValues.Back) return true;
        if (value == PageBorderZOrderValues.Front) return false;
        return null;
    }

    private static PageBorderZOrderValues? ToWordPageBorderZOrder(bool? displayBehindText) {
        if (!displayBehindText.HasValue) return null;
        return displayBehindText.Value ? PageBorderZOrderValues.Back : PageBorderZOrderValues.Front;
    }

    private static RtfPageBorderOffset? ToRtfPageBorderOffset(PageBorderOffsetValues? value) {
        if (value == PageBorderOffsetValues.Page) return RtfPageBorderOffset.PageEdge;
        if (value == PageBorderOffsetValues.Text) return RtfPageBorderOffset.Text;
        return null;
    }

    private static PageBorderOffsetValues? ToWordPageBorderOffset(RtfPageBorderOffset? value) {
        if (value == RtfPageBorderOffset.PageEdge) return PageBorderOffsetValues.Page;
        if (value == RtfPageBorderOffset.Text) return PageBorderOffsetValues.Text;
        return null;
    }

    private static void CopyNoteSettings(FootnoteProperties? footnoteProperties, EndnoteProperties? endnoteProperties, RtfNoteSettings destination) {
        if (footnoteProperties != null) {
            destination.FootnoteStartNumber = ToInt32(footnoteProperties.NumberingStart?.Val);
            destination.FootnoteRestart = ToRtfNoteRestart(footnoteProperties.NumberingRestart?.Val?.Value);
            destination.FootnoteNumberFormat = ToRtfNoteNumberFormat(footnoteProperties.NumberingFormat?.Val?.Value);
            destination.FootnotePlacement = ToRtfFootnotePlacement(footnoteProperties.FootnotePosition?.Val?.Value);
        }

        if (endnoteProperties != null) {
            destination.EndnoteStartNumber = ToInt32(endnoteProperties.NumberingStart?.Val);
            destination.EndnoteRestart = ToRtfNoteRestart(endnoteProperties.NumberingRestart?.Val?.Value);
            destination.EndnoteNumberFormat = ToRtfNoteNumberFormat(endnoteProperties.NumberingFormat?.Val?.Value);
            destination.EndnotePlacement = ToRtfEndnotePlacement(endnoteProperties.EndnotePosition?.Val?.Value);
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

    private static RtfNoteNumberFormat? ToRtfNoteNumberFormat(NumberFormatValues? format) {
        if (format == NumberFormatValues.LowerRoman) return RtfNoteNumberFormat.LowerRoman;
        if (format == NumberFormatValues.UpperRoman) return RtfNoteNumberFormat.UpperRoman;
        if (format == NumberFormatValues.LowerLetter) return RtfNoteNumberFormat.LowerLetter;
        if (format == NumberFormatValues.UpperLetter) return RtfNoteNumberFormat.UpperLetter;
        if (format == NumberFormatValues.Decimal) return RtfNoteNumberFormat.Arabic;
        return null;
    }

    private static NumberFormatValues? ToWordNoteNumberFormat(RtfNoteNumberFormat? format) {
        if (format == RtfNoteNumberFormat.LowerRoman) return NumberFormatValues.LowerRoman;
        if (format == RtfNoteNumberFormat.UpperRoman) return NumberFormatValues.UpperRoman;
        if (format == RtfNoteNumberFormat.LowerLetter) return NumberFormatValues.LowerLetter;
        if (format == RtfNoteNumberFormat.UpperLetter) return NumberFormatValues.UpperLetter;
        if (format == RtfNoteNumberFormat.Arabic) return NumberFormatValues.Decimal;
        return null;
    }

    private static RtfNoteNumberRestart? ToRtfNoteRestart(RestartNumberValues? restart) {
        if (restart == RestartNumberValues.EachPage) return RtfNoteNumberRestart.EachPage;
        if (restart == RestartNumberValues.EachSection) return RtfNoteNumberRestart.EachSection;
        if (restart == RestartNumberValues.Continuous) return RtfNoteNumberRestart.Continuous;
        return null;
    }

    private static RestartNumberValues? ToWordNoteRestart(RtfNoteNumberRestart? restart) {
        if (restart == RtfNoteNumberRestart.EachPage) return RestartNumberValues.EachPage;
        if (restart == RtfNoteNumberRestart.EachSection) return RestartNumberValues.EachSection;
        if (restart == RtfNoteNumberRestart.Continuous) return RestartNumberValues.Continuous;
        return null;
    }

    private static void CopyLineNumbering(LineNumberType? source, RtfLineNumbering destination) {
        if (source == null) {
            return;
        }

        destination.CountBy = ToInt32(source.CountBy);
        destination.StartNumber = ToInt32(source.Start);
        destination.DistanceFromTextTwips = ToInt32(source.Distance);
        destination.Restart = ToRtfLineNumberRestart(source.Restart?.Value);
    }

    private static void ApplyLineNumbering(RtfLineNumbering source, WordSection destination) {
        if (!source.HasAnyValue) {
            return;
        }

        var lineNumberType = new LineNumberType {
            Restart = ToWordLineNumberRestart(source.Restart)
        };
        if (source.CountBy.HasValue) {
            lineNumberType.CountBy = ToInt16Value(source.CountBy.Value, nameof(source.CountBy));
        }

        if (source.StartNumber.HasValue) {
            lineNumberType.Start = ToInt16Value(source.StartNumber.Value, nameof(source.StartNumber));
        }

        if (source.DistanceFromTextTwips.HasValue) {
            lineNumberType.Distance = source.DistanceFromTextTwips.Value.ToString(CultureInfo.InvariantCulture);
        }

        destination._sectionProperties.RemoveAllChildren<LineNumberType>();
        destination._sectionProperties.Append(lineNumberType);
    }

    private static RtfLineNumberRestart? ToRtfLineNumberRestart(LineNumberRestartValues? restart) {
        if (restart == LineNumberRestartValues.NewPage) return RtfLineNumberRestart.EachPage;
        if (restart == LineNumberRestartValues.NewSection) return RtfLineNumberRestart.EachSection;
        if (restart == LineNumberRestartValues.Continuous) return RtfLineNumberRestart.Continuous;
        return null;
    }

    private static LineNumberRestartValues? ToWordLineNumberRestart(RtfLineNumberRestart? restart) {
        if (restart == RtfLineNumberRestart.EachPage) return LineNumberRestartValues.NewPage;
        if (restart == RtfLineNumberRestart.EachSection) return LineNumberRestartValues.NewSection;
        if (restart == RtfLineNumberRestart.Continuous) return LineNumberRestartValues.Continuous;
        return null;
    }

    private static RtfSectionVerticalAlignment? ToRtfSectionVerticalAlignment(VerticalJustificationValues? alignment) {
        if (alignment == VerticalJustificationValues.Top) return RtfSectionVerticalAlignment.Top;
        if (alignment == VerticalJustificationValues.Center) return RtfSectionVerticalAlignment.Center;
        if (alignment == VerticalJustificationValues.Bottom) return RtfSectionVerticalAlignment.Bottom;
        if (alignment == VerticalJustificationValues.Both) return RtfSectionVerticalAlignment.Justified;
        return null;
    }

    private static void ApplySectionVerticalAlignment(RtfSectionVerticalAlignment? alignment, WordSection destination) {
        VerticalJustificationValues? wordAlignment = ToWordSectionVerticalAlignment(alignment);
        if (!wordAlignment.HasValue) {
            return;
        }

        destination._sectionProperties.RemoveAllChildren<VerticalTextAlignmentOnPage>();
        destination._sectionProperties.Append(new VerticalTextAlignmentOnPage { Val = wordAlignment.Value });
    }

    private static VerticalJustificationValues? ToWordSectionVerticalAlignment(RtfSectionVerticalAlignment? alignment) {
        if (alignment == RtfSectionVerticalAlignment.Top) return VerticalJustificationValues.Top;
        if (alignment == RtfSectionVerticalAlignment.Center) return VerticalJustificationValues.Center;
        if (alignment == RtfSectionVerticalAlignment.Bottom) return VerticalJustificationValues.Bottom;
        if (alignment == RtfSectionVerticalAlignment.Justified) return VerticalJustificationValues.Both;
        return null;
    }

    private static RtfFootnotePlacement? ToRtfFootnotePlacement(FootnotePositionValues? placement) {
        if (placement == FootnotePositionValues.PageBottom) return RtfFootnotePlacement.PageBottom;
        if (placement == FootnotePositionValues.BeneathText) return RtfFootnotePlacement.BeneathText;
        if (placement == FootnotePositionValues.SectionEnd) return RtfFootnotePlacement.SectionEnd;
        return null;
    }

    private static FootnotePositionValues? ToWordFootnotePlacement(RtfFootnotePlacement? placement) {
        if (placement == RtfFootnotePlacement.PageBottom) return FootnotePositionValues.PageBottom;
        if (placement == RtfFootnotePlacement.BeneathText) return FootnotePositionValues.BeneathText;
        if (placement == RtfFootnotePlacement.SectionEnd) return FootnotePositionValues.SectionEnd;
        return null;
    }

    private static RtfEndnotePlacement? ToRtfEndnotePlacement(EndnotePositionValues? placement) {
        if (placement == EndnotePositionValues.SectionEnd) return RtfEndnotePlacement.SectionEnd;
        if (placement == EndnotePositionValues.DocumentEnd) return RtfEndnotePlacement.DocumentEnd;
        return null;
    }

    private static EndnotePositionValues? ToWordEndnotePlacement(RtfEndnotePlacement? placement) {
        if (placement == RtfEndnotePlacement.SectionEnd) return EndnotePositionValues.SectionEnd;
        if (placement == RtfEndnotePlacement.DocumentEnd) return EndnotePositionValues.DocumentEnd;
        return null;
    }

}