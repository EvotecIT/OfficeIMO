using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;

#pragma warning disable RS0026
namespace OfficeIMO.Excel {
    /// <summary>
    /// Page setup of a worksheet: paper, orientation, scaling, margins, print area and titles, print options,
    /// and headers and footers. Values are read from and written to the worksheet directly.
    /// </summary>
    public sealed class ExcelSheetPageSetup {
        private readonly ExcelSheet _sheet;

        internal ExcelSheetPageSetup(ExcelSheet sheet) {
            _sheet = sheet;
            Header = new ExcelHeaderFooter(sheet, ExcelHeaderFooterKind.OddHeader);
            Footer = new ExcelHeaderFooter(sheet, ExcelHeaderFooterKind.OddFooter);
            FirstPageHeader = new ExcelHeaderFooter(sheet, ExcelHeaderFooterKind.FirstHeader);
            FirstPageFooter = new ExcelHeaderFooter(sheet, ExcelHeaderFooterKind.FirstFooter);
            EvenPageHeader = new ExcelHeaderFooter(sheet, ExcelHeaderFooterKind.EvenHeader);
            EvenPageFooter = new ExcelHeaderFooter(sheet, ExcelHeaderFooterKind.EvenFooter);
        }

        private PageSetup? Element => _sheet.WorksheetRoot.GetFirstChild<PageSetup>();

        /// <summary>Page orientation; null when the worksheet does not store one (Excel prints portrait).</summary>
        public OfficePageOrientation? Orientation {
            get {
                OrientationValues? value = Element?.Orientation?.Value;
                if (value == OrientationValues.Landscape) return OfficePageOrientation.Landscape;
                if (value == OrientationValues.Portrait) return OfficePageOrientation.Portrait;
                return null;
            }
            set {
                if (value.HasValue) _sheet.SetOrientationCore(value.Value);
                else _sheet.ClearPageSetupAttribute("orientation");
            }
        }

        /// <summary>Known paper size; null when none is stored or the stored code has no <see cref="ExcelPaperSize"/> value.</summary>
        public ExcelPaperSize? PaperSize {
            get => Element?.PaperSize?.Value is uint code && Enum.IsDefined(typeof(ExcelPaperSize), code) ? (ExcelPaperSize)code : null;
            set {
                if (value.HasValue) _sheet.SetPaperSizeCore(value.Value);
                else _sheet.ClearPageSetupAttribute("paperSize");
            }
        }

        /// <summary>Stored OpenXML paper-size code, including codes without an <see cref="ExcelPaperSize"/> value.</summary>
        public uint? PaperSizeCode => Element?.PaperSize?.Value;

        /// <summary>
        /// Pages to fit horizontally when fit-to-page is on; null when it is off. Setting a value turns fit-to-page on.
        /// An omitted dimension defaults to one; zero leaves the axis unlimited.
        /// </summary>
        public uint? FitToWidth {
            get => IsFitToPage ? Element?.FitToWidth?.Value ?? 1U : null;
            set => _sheet.SetFitToPageDimension(width: true, value);
        }

        /// <summary>
        /// Pages to fit vertically when fit-to-page is on; null when it is off. Setting a value turns fit-to-page on.
        /// An omitted dimension defaults to one; zero leaves the axis unlimited.
        /// </summary>
        public uint? FitToHeight {
            get => IsFitToPage ? Element?.FitToHeight?.Value ?? 1U : null;
            set => _sheet.SetFitToPageDimension(width: false, value);
        }

        /// <summary>Manual print scale percentage (10-400); null when none is stored.</summary>
        public uint? Scale {
            get => Element?.Scale?.Value;
            set {
                if (value.HasValue) _sheet.SetPageSetupCore(scale: value.Value);
                else _sheet.ClearPageSetupAttribute("scale");
            }
        }

        /// <summary>Order in which multi-page sheets print; null when none is stored.</summary>
        public ExcelPageOrder? PageOrder {
            get {
                PageOrderValues? value = Element?.PageOrder?.Value;
                if (value == PageOrderValues.OverThenDown) return ExcelPageOrder.OverThenDown;
                if (value == PageOrderValues.DownThenOver) return ExcelPageOrder.DownThenOver;
                return null;
            }
            set {
                if (value.HasValue) _sheet.SetPageSetupCore(pageOrder: value.Value);
                else _sheet.ClearPageSetupAttribute("pageOrder");
            }
        }

        private bool IsFitToPage => _sheet.WorksheetRoot.GetFirstChild<SheetProperties>()?
            .GetFirstChild<PageSetupProperties>()?.FitToPage?.Value == true;

        /// <summary>Page margins; null when the worksheet does not store them.</summary>
        public ExcelSheetPageMargins? Margins {
            get {
                PageMargins? margins = _sheet.WorksheetRoot.GetFirstChild<PageMargins>();
                return margins == null ? null : new ExcelSheetPageMargins {
                    Left = OfficeLength.Inches(margins.Left?.Value ?? 0D),
                    Right = OfficeLength.Inches(margins.Right?.Value ?? 0D),
                    Top = OfficeLength.Inches(margins.Top?.Value ?? 0D),
                    Bottom = OfficeLength.Inches(margins.Bottom?.Value ?? 0D),
                    Header = OfficeLength.Inches(margins.Header?.Value ?? 0D),
                    Footer = OfficeLength.Inches(margins.Footer?.Value ?? 0D)
                };
            }
        }

        /// <summary>Sets physical page margins. Omitted header and footer margins are 0.3 inches.</summary>
        public ExcelSheetPageSetup SetMargins(OfficeLength left, OfficeLength right, OfficeLength top, OfficeLength bottom, OfficeLength? header = null, OfficeLength? footer = null) {
            _sheet.SetMarginsCore(left, right, top, bottom, header, footer);
            return this;
        }

        /// <summary>Sets the margins of one of Excel's margin presets.</summary>
        public ExcelSheetPageSetup SetMargins(ExcelMarginPreset preset) {
            _sheet.SetMarginsPresetCore(preset);
            return this;
        }

        /// <summary>
        /// Print area as stored, with one or more comma-separated areas; null when none is set. Setting null removes it.
        /// </summary>
        public string? PrintArea {
            get => _sheet.GetPrintAreaText();
            set {
                if (value == null) _sheet.Document.RemovePrintArea(_sheet);
                else _sheet.Document.SetPrintAreaCore(_sheet, value);
            }
        }

        /// <summary>The separate print areas in their stored order; empty when none is set.</summary>
        public IReadOnlyList<string> PrintAreas => _sheet.GetPrintAreaParts();

        /// <summary>Rows and columns repeated on every printed page.</summary>
        public ExcelPrintTitles PrintTitles => _sheet.ReadPrintTitles();

        /// <summary>Sets the rows and columns repeated on every printed page; nulls on both axes clear them.</summary>
        public ExcelSheetPageSetup SetPrintTitles(int? firstRow, int? lastRow, int? firstColumn = null, int? lastColumn = null) {
            _sheet.Document.SetPrintTitlesCore(_sheet, firstRow, lastRow, firstColumn, lastColumn);
            return this;
        }

        /// <summary>Whether cell gridlines print.</summary>
        public bool PrintGridLines {
            get => _sheet.ReadPrintOptions().PrintGridLines ?? false;
            set => _sheet.SetPrintOptionsCore(printGridLines: value);
        }

        /// <summary>Whether row and column headings print.</summary>
        public bool PrintHeadings {
            get => _sheet.ReadPrintOptions().PrintHeadings ?? false;
            set => _sheet.SetPrintOptionsCore(printHeadings: value);
        }

        /// <summary>Whether the printed sheet is centered horizontally on the page.</summary>
        public bool HorizontalCentered {
            get => _sheet.ReadPrintOptions().HorizontalCentered ?? false;
            set => _sheet.SetPrintOptionsCore(horizontalCentered: value);
        }

        /// <summary>Whether the printed sheet is centered vertically on the page.</summary>
        public bool VerticalCentered {
            get => _sheet.ReadPrintOptions().VerticalCentered ?? false;
            set => _sheet.SetPrintOptionsCore(verticalCentered: value);
        }

        /// <summary>Applies a print layout preset, such as Report or Wide, with optional overrides.</summary>
        public ExcelSheetPageSetup ApplyLayout(ExcelPrintLayoutOptions options) {
            _sheet.ApplyPrintLayoutCore(options);
            return this;
        }

        /// <summary>Applies a print layout preset and an optional print area.</summary>
        public ExcelSheetPageSetup ApplyLayout(ExcelPrintLayoutPreset preset, string? printArea = null) {
            _sheet.ApplyPrintLayoutPresetCore(preset, printArea);
            return this;
        }

        /// <summary>Header on every page, or on odd pages when <see cref="DifferentOddEven"/> is on.</summary>
        public ExcelHeaderFooter Header { get; }

        /// <summary>Footer on every page, or on odd pages when <see cref="DifferentOddEven"/> is on.</summary>
        public ExcelHeaderFooter Footer { get; }

        /// <summary>Header on the first page; writing text turns <see cref="DifferentFirstPage"/> on.</summary>
        public ExcelHeaderFooter FirstPageHeader { get; }

        /// <summary>Footer on the first page; writing text turns <see cref="DifferentFirstPage"/> on.</summary>
        public ExcelHeaderFooter FirstPageFooter { get; }

        /// <summary>Header on even pages; writing text turns <see cref="DifferentOddEven"/> on.</summary>
        public ExcelHeaderFooter EvenPageHeader { get; }

        /// <summary>Footer on even pages; writing text turns <see cref="DifferentOddEven"/> on.</summary>
        public ExcelHeaderFooter EvenPageFooter { get; }

        /// <summary>
        /// Whether the first page uses its own header and footer. Turning it on copies the running header and
        /// footer into a missing first-page variant; existing first-page text is kept when it is turned off.
        /// </summary>
        public bool DifferentFirstPage {
            get => HeaderFooterElement?.DifferentFirst?.Value ?? false;
            set => _sheet.SetHeaderCore(differentFirstPage: value);
        }

        /// <summary>
        /// Whether even pages use their own header and footer. Turning it on copies the running header and footer
        /// into a missing even-page variant; existing even-page text is kept when it is turned off.
        /// </summary>
        public bool DifferentOddEven {
            get => HeaderFooterElement?.DifferentOddEven?.Value ?? false;
            set => _sheet.SetHeaderCore(differentOddEven: value);
        }

        /// <summary>Whether headers and footers align with the page margins.</summary>
        public bool AlignHeaderFooterWithMargins {
            get => HeaderFooterElement?.AlignWithMargins?.Value ?? true;
            set => _sheet.SetHeaderCore(alignWithMargins: value);
        }

        /// <summary>Whether headers and footers scale with the document.</summary>
        public bool ScaleHeaderFooterWithDocument {
            get => HeaderFooterElement?.ScaleWithDoc?.Value ?? true;
            set => _sheet.SetHeaderCore(scaleWithDocument: value);
        }

        private HeaderFooter? HeaderFooterElement => _sheet.WorksheetRoot.GetFirstChild<HeaderFooter>();
    }

    internal enum ExcelHeaderFooterKind {
        OddHeader,
        OddFooter,
        FirstHeader,
        FirstFooter,
        EvenHeader,
        EvenFooter
    }

    /// <summary>One header or footer of a worksheet, with left, center and right sections.</summary>
    public sealed class ExcelHeaderFooter {
        internal ExcelHeaderFooter(ExcelSheet sheet, ExcelHeaderFooterKind kind) {
            Left = new ExcelHeaderFooterSection(sheet, kind, ExcelHeaderFooterPosition.Left);
            Center = new ExcelHeaderFooterSection(sheet, kind, ExcelHeaderFooterPosition.Center);
            Right = new ExcelHeaderFooterSection(sheet, kind, ExcelHeaderFooterPosition.Right);
        }

        /// <summary>Left section.</summary>
        public ExcelHeaderFooterSection Left { get; }

        /// <summary>Center section.</summary>
        public ExcelHeaderFooterSection Center { get; }

        /// <summary>Right section.</summary>
        public ExcelHeaderFooterSection Right { get; }
    }

    /// <summary>A left, center or right section of a header or footer: text with Excel codes, and an optional picture.</summary>
    public sealed class ExcelHeaderFooterSection {
        private readonly ExcelSheet _sheet;
        private readonly ExcelHeaderFooterKind _kind;

        internal ExcelHeaderFooterSection(ExcelSheet sheet, ExcelHeaderFooterKind kind, ExcelHeaderFooterPosition position) {
            _sheet = sheet;
            _kind = kind;
            Position = position;
        }

        /// <summary>Which section this is.</summary>
        public ExcelHeaderFooterPosition Position { get; }

        /// <summary>
        /// Section text. Excel codes such as &amp;P (page), &amp;N (pages), &amp;D (date), &amp;A (sheet name) and
        /// &amp;G (picture) are kept; other ampersands are written literally. An empty string clears the section.
        /// </summary>
        public string Text {
            get => _sheet.GetHeaderFooterSectionText(_kind, Position);
            set => _sheet.SetHeaderFooterSectionText(_kind, Position, value ?? string.Empty);
        }

        /// <summary>The section picture; null when it has none. Pictures are supported in the running header and footer.</summary>
        public ExcelHeaderFooterImageInfo? Image => _sheet.GetHeaderFooterSectionImage(_kind, Position);

        /// <summary>Sets the section picture and adds the &amp;G placeholder to its text.</summary>
        /// <param name="imageBytes">Image bytes.</param>
        /// <param name="contentType">Image content type, such as image/png or image/jpeg.</param>
        /// <param name="width">Optional physical width; inferred from the image when omitted.</param>
        /// <param name="height">Optional physical height; inferred from the image when omitted.</param>
        /// <exception cref="NotSupportedException">The section belongs to a first-page or even-page variant.</exception>
        public void SetImage(byte[] imageBytes, string contentType = "image/png", OfficeLength? width = null, OfficeLength? height = null) {
            if (RequireRunningHeaderOrFooter()) _sheet.SetHeaderImageCore(Position, imageBytes, contentType, width, height);
            else _sheet.SetFooterImageCore(Position, imageBytes, contentType, width, height);
        }

        /// <summary>Downloads an image and sets it as the section picture.</summary>
        public Task SetImageFromUrlAsync(string url, OfficeLength? width = null, OfficeLength? height = null, CancellationToken cancellationToken = default) =>
            SetImageFromUrlCoreAsync(url, null, width, height, cancellationToken);

        /// <summary>Downloads an image under an explicit remote-network policy and sets it as the section picture.</summary>
        public Task SetImageFromUrlAsync(string url, OfficeRemoteImageLoadOptions remoteImageOptions, OfficeLength? width = null,
            OfficeLength? height = null, CancellationToken cancellationToken = default) {
            if (remoteImageOptions == null) throw new ArgumentNullException(nameof(remoteImageOptions));
            return SetImageFromUrlCoreAsync(url, remoteImageOptions, width, height, cancellationToken);
        }

        private Task SetImageFromUrlCoreAsync(string url, OfficeRemoteImageLoadOptions? options, OfficeLength? width, OfficeLength? height, CancellationToken cancellationToken) {
            return RequireRunningHeaderOrFooter()
                ? _sheet.SetHeaderImageFromUrlCoreAsync(Position, url, options, width, height, cancellationToken)
                : _sheet.SetFooterImageFromUrlCoreAsync(Position, url, options, width, height, cancellationToken);
        }

        /// <summary>Returns true for the running header and false for the running footer.</summary>
        private bool RequireRunningHeaderOrFooter() {
            if (_kind == ExcelHeaderFooterKind.OddHeader) return true;
            if (_kind == ExcelHeaderFooterKind.OddFooter) return false;
            throw new NotSupportedException("Pictures are supported in the running header and footer only.");
        }
    }
}
#pragma warning restore RS0026
