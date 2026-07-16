namespace OfficeIMO.Excel.Xlsb.Model {
    /// <summary>Represents the six standard worksheet page margins stored in BIFF12.</summary>
    internal sealed class XlsbPageMargins {
        internal XlsbPageMargins(double left, double right, double top, double bottom, double header, double footer) {
            Left = left;
            Right = right;
            Top = top;
            Bottom = bottom;
            Header = header;
            Footer = footer;
        }

        internal double Left { get; }
        internal double Right { get; }
        internal double Top { get; }
        internal double Bottom { get; }
        internal double Header { get; }
        internal double Footer { get; }
    }

    /// <summary>Represents the normal worksheet print-option flags stored in BIFF12.</summary>
    internal sealed class XlsbPrintOptions {
        internal XlsbPrintOptions(bool horizontalCentered, bool verticalCentered, bool headings, bool gridLines) {
            HorizontalCentered = horizontalCentered;
            VerticalCentered = verticalCentered;
            Headings = headings;
            GridLines = gridLines;
        }

        internal bool HorizontalCentered { get; }
        internal bool VerticalCentered { get; }
        internal bool Headings { get; }
        internal bool GridLines { get; }
    }

    /// <summary>Identifies how worksheet cell errors are rendered when printed.</summary>
    internal enum XlsbPrintErrorMode {
        Displayed = 0,
        Blank = 1,
        Dash = 2,
        NotAvailable = 3
    }

    /// <summary>Represents the standard worksheet page-setup fields stored in BIFF12.</summary>
    internal sealed class XlsbPageSetup {
        internal XlsbPageSetup(
            uint paperSize,
            uint scale,
            uint horizontalDpi,
            uint verticalDpi,
            uint copies,
            int firstPageNumber,
            uint fitToWidth,
            uint fitToHeight,
            bool overThenDown,
            bool landscape,
            bool blackAndWhite,
            bool draft,
            bool printCellComments,
            bool useDefaultOrientation,
            bool useFirstPageNumber,
            bool commentsAtEnd,
            XlsbPrintErrorMode errors,
            string? printerSettingsRelationshipId) {
            PaperSize = paperSize;
            Scale = scale;
            HorizontalDpi = horizontalDpi;
            VerticalDpi = verticalDpi;
            Copies = copies;
            FirstPageNumber = firstPageNumber;
            FitToWidth = fitToWidth;
            FitToHeight = fitToHeight;
            OverThenDown = overThenDown;
            Landscape = landscape;
            BlackAndWhite = blackAndWhite;
            Draft = draft;
            PrintCellComments = printCellComments;
            UseDefaultOrientation = useDefaultOrientation;
            UseFirstPageNumber = useFirstPageNumber;
            CommentsAtEnd = commentsAtEnd;
            Errors = errors;
            PrinterSettingsRelationshipId = printerSettingsRelationshipId;
        }

        internal uint PaperSize { get; }
        internal uint Scale { get; }
        internal uint HorizontalDpi { get; }
        internal uint VerticalDpi { get; }
        internal uint Copies { get; }
        internal int FirstPageNumber { get; }
        internal uint FitToWidth { get; }
        internal uint FitToHeight { get; }
        internal bool OverThenDown { get; }
        internal bool Landscape { get; }
        internal bool BlackAndWhite { get; }
        internal bool Draft { get; }
        internal bool PrintCellComments { get; }
        internal bool UseDefaultOrientation { get; }
        internal bool UseFirstPageNumber { get; }
        internal bool CommentsAtEnd { get; }
        internal XlsbPrintErrorMode Errors { get; }
        internal string? PrinterSettingsRelationshipId { get; }
    }

    /// <summary>Represents textual worksheet headers and footers stored in BIFF12.</summary>
    internal sealed class XlsbHeaderFooter {
        internal XlsbHeaderFooter(
            bool differentOddEven,
            bool differentFirst,
            bool scaleWithDocument,
            bool alignWithMargins,
            string? oddHeader,
            string? oddFooter,
            string? evenHeader,
            string? evenFooter,
            string? firstHeader,
            string? firstFooter) {
            DifferentOddEven = differentOddEven;
            DifferentFirst = differentFirst;
            ScaleWithDocument = scaleWithDocument;
            AlignWithMargins = alignWithMargins;
            OddHeader = oddHeader;
            OddFooter = oddFooter;
            EvenHeader = evenHeader;
            EvenFooter = evenFooter;
            FirstHeader = firstHeader;
            FirstFooter = firstFooter;
        }

        internal bool DifferentOddEven { get; }
        internal bool DifferentFirst { get; }
        internal bool ScaleWithDocument { get; }
        internal bool AlignWithMargins { get; }
        internal string? OddHeader { get; }
        internal string? OddFooter { get; }
        internal string? EvenHeader { get; }
        internal string? EvenFooter { get; }
        internal string? FirstHeader { get; }
        internal string? FirstFooter { get; }
    }
}
