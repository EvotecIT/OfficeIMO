using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

internal static partial class RtfSemanticReader {
    private sealed partial class Binder {
        private bool TryApplySectionControl(RtfControlWord control, CharacterState state) {
            switch (control.Name) {
                case "sect":
                    FlushParagraphIfNeeded(force: false, state);
                    _hasSemanticSections = true;
                    CompleteCurrentSection();
                    return true;
                case "sectd":
                    _hasSemanticSections = true;
                    EnsureCurrentSection().ResetLayout();
                    _currentSectionColumnNumber = null;
                    state.CurrentPageBorderSide = null;
                    return true;
                case "sbknone":
                    SetSectionBreakKind(RtfSectionBreakKind.Continuous);
                    return true;
                case "sbkcol":
                    SetSectionBreakKind(RtfSectionBreakKind.Column);
                    return true;
                case "sbkpage":
                    SetSectionBreakKind(RtfSectionBreakKind.NextPage);
                    return true;
                case "sbkeven":
                    SetSectionBreakKind(RtfSectionBreakKind.EvenPage);
                    return true;
                case "sbkodd":
                    SetSectionBreakKind(RtfSectionBreakKind.OddPage);
                    return true;
                case "cols":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().ColumnCount = control.Parameter;
                    return true;
                case "colsx":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().ColumnSpaceTwips = control.Parameter;
                    return true;
                case "linebetcol":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().ColumnSeparator = !control.HasParameter || control.Parameter != 0;
                    return true;
                case "colno":
                    if (!_hasSemanticSections) return false;
                    _currentSectionColumnNumber = control.Parameter.GetValueOrDefault(EnsureCurrentSection().Columns.Count + 1);
                    EnsureCurrentSection().EnsureColumn(_currentSectionColumnNumber.Value);
                    return true;
                case "colw":
                    if (!_hasSemanticSections) return false;
                    GetCurrentSectionColumn().WidthTwips = control.Parameter;
                    return true;
                case "colsr":
                    if (!_hasSemanticSections) return false;
                    GetCurrentSectionColumn().SpaceAfterTwips = control.Parameter;
                    return true;
                case "linemod":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().LineNumbering.CountBy = control.Parameter;
                    return true;
                case "linex":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().LineNumbering.DistanceFromTextTwips = control.Parameter;
                    return true;
                case "linestarts":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().LineNumbering.StartNumber = control.Parameter;
                    return true;
                case "linerestart":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().LineNumbering.Restart = RtfLineNumberRestart.EachSection;
                    return true;
                case "lineppage":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().LineNumbering.Restart = RtfLineNumberRestart.EachPage;
                    return true;
                case "linecont":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().LineNumbering.Restart = RtfLineNumberRestart.Continuous;
                    return true;
                case "vertalt":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().VerticalAlignment = RtfSectionVerticalAlignment.Top;
                    return true;
                case "vertalc":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().VerticalAlignment = RtfSectionVerticalAlignment.Center;
                    return true;
                case "vertalb":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().VerticalAlignment = RtfSectionVerticalAlignment.Bottom;
                    return true;
                case "vertalj":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().VerticalAlignment = RtfSectionVerticalAlignment.Justified;
                    return true;
                case "ltrsect":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().Direction = RtfTextDirection.LeftToRight;
                    return true;
                case "rtlsect":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().Direction = RtfTextDirection.RightToLeft;
                    return true;
                case "paperw":
                case "pgwsxn":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().PageSetup.PaperWidthTwips = control.Parameter;
                    return true;
                case "paperh":
                case "pghsxn":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().PageSetup.PaperHeightTwips = control.Parameter;
                    return true;
                case "psz":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().PageSetup.PrinterPaperSize = control.Parameter;
                    return true;
                case "binfsxn":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().PageSetup.FirstPagePaperSource = control.Parameter;
                    return true;
                case "binsxn":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().PageSetup.OtherPagesPaperSource = control.Parameter;
                    return true;
                case "margl":
                case "marglsxn":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().PageSetup.MarginLeftTwips = control.Parameter;
                    return true;
                case "margr":
                case "margrsxn":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().PageSetup.MarginRightTwips = control.Parameter;
                    return true;
                case "margt":
                case "margtsxn":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().PageSetup.MarginTopTwips = control.Parameter;
                    return true;
                case "margb":
                case "margbsxn":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().PageSetup.MarginBottomTwips = control.Parameter;
                    return true;
                case "gutter":
                case "guttersxn":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().PageSetup.GutterWidthTwips = control.Parameter;
                    return true;
                case "headery":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().PageSetup.HeaderDistanceTwips = control.Parameter;
                    return true;
                case "footery":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().PageSetup.FooterDistanceTwips = control.Parameter;
                    return true;
                case "rtlgutter":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().PageSetup.RtlGutter = !control.HasParameter || control.Parameter != 0;
                    return true;
                case "pgnstarts":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().PageSetup.PageNumberStart = control.Parameter;
                    return true;
                case "pgncont":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().PageSetup.PageNumberRestart = false;
                    return true;
                case "pgnrestart":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().PageSetup.PageNumberRestart = true;
                    return true;
                case "pgnx":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().PageSetup.PageNumberPositionXTwips = control.Parameter;
                    return true;
                case "pgny":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().PageSetup.PageNumberPositionYTwips = control.Parameter;
                    return true;
                case "pgndec":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().PageSetup.PageNumberFormat = RtfPageNumberFormat.Decimal;
                    return true;
                case "pgnucrm":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().PageSetup.PageNumberFormat = RtfPageNumberFormat.UpperRoman;
                    return true;
                case "pgnlcrm":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().PageSetup.PageNumberFormat = RtfPageNumberFormat.LowerRoman;
                    return true;
                case "pgnucltr":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().PageSetup.PageNumberFormat = RtfPageNumberFormat.UpperLetter;
                    return true;
                case "pgnlcltr":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().PageSetup.PageNumberFormat = RtfPageNumberFormat.LowerLetter;
                    return true;
                case "pgndecd":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().PageSetup.PageNumberFormat = RtfPageNumberFormat.DoubleByteDecimal;
                    return true;
                case "landscape":
                case "lndscpsxn":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().PageSetup.Landscape = !control.HasParameter || control.Parameter != 0;
                    return true;
                case "titlepg":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().PageSetup.DifferentFirstPageHeaderFooter = !control.HasParameter || control.Parameter != 0;
                    return true;
                case "pgbrdrhead":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().PageSetup.PageBorders.IncludeHeader = !control.HasParameter || control.Parameter != 0;
                    state.CurrentPageBorderSide = null;
                    return true;
                case "pgbrdrfoot":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().PageSetup.PageBorders.IncludeFooter = !control.HasParameter || control.Parameter != 0;
                    state.CurrentPageBorderSide = null;
                    return true;
                case "pgbrdrsnap":
                    if (!_hasSemanticSections) return false;
                    EnsureCurrentSection().PageSetup.PageBorders.SnapToPageBorder = !control.HasParameter || control.Parameter != 0;
                    state.CurrentPageBorderSide = null;
                    return true;
                case "pgbrdropt":
                    if (!_hasSemanticSections) return false;
                    ApplyPageBorderDisplayOptions(EnsureCurrentSection().PageSetup.PageBorders, control.Parameter.GetValueOrDefault());
                    state.CurrentPageBorderSide = null;
                    return true;
                case "pgbrdrt":
                    if (!_hasSemanticSections) return false;
                    BeginPageBorder(state, RtfPageBorderSide.Top);
                    return true;
                case "pgbrdrb":
                    if (!_hasSemanticSections) return false;
                    BeginPageBorder(state, RtfPageBorderSide.Bottom);
                    return true;
                case "pgbrdrl":
                    if (!_hasSemanticSections) return false;
                    BeginPageBorder(state, RtfPageBorderSide.Left);
                    return true;
                case "pgbrdrr":
                    if (!_hasSemanticSections) return false;
                    BeginPageBorder(state, RtfPageBorderSide.Right);
                    return true;
                case "brdrs":
                    return TryApplyCurrentPageBorderStyle(state, RtfPageBorderStyle.Single);
                case "brdrdb":
                    return TryApplyCurrentPageBorderStyle(state, RtfPageBorderStyle.Double);
                case "brdrdot":
                    return TryApplyCurrentPageBorderStyle(state, RtfPageBorderStyle.Dotted);
                case "brdrdash":
                    return TryApplyCurrentPageBorderStyle(state, RtfPageBorderStyle.Dashed);
                case "brdrsh":
                    return TryApplyCurrentPageBorderStyle(state, RtfPageBorderStyle.Shadow, shadow: true);
                case "brdrnone":
                case "brdrnil":
                    return TryApplyCurrentPageBorderStyle(state, RtfPageBorderStyle.None);
                case "brdrw":
                    return TryApplyCurrentPageBorderWidth(state, control.Parameter);
                case "brsp":
                    return TryApplyCurrentPageBorderSpace(state, control.Parameter);
                case "brdrcf":
                    return TryApplyCurrentPageBorderColor(state, control.Parameter);
                case "brdrframe":
                    return TryApplyCurrentPageBorderFrame(state, !control.HasParameter || control.Parameter != 0);
                default:
                    if (!_hasSemanticSections) return false;
                    return TryApplyNoteSettingsControl(control, EnsureCurrentSection().NoteSettings);
            }
        }

        private void SetSectionBreakKind(RtfSectionBreakKind kind) {
            _hasSemanticSections = true;
            EnsureCurrentSection().BreakKind = kind;
        }

        private RtfSection EnsureCurrentSection() {
            return _currentSection ??= new RtfSection();
        }

        private RtfSectionColumn GetCurrentSectionColumn() {
            int oneBasedIndex = _currentSectionColumnNumber.GetValueOrDefault(EnsureCurrentSection().Columns.Count + 1);
            RtfSectionColumn column = EnsureCurrentSection().EnsureColumn(oneBasedIndex);
            _currentSectionColumnNumber = oneBasedIndex;
            return column;
        }

        private void BeginPageBorder(CharacterState state, RtfPageBorderSide side) {
            state.CurrentPageBorderSide = side;
            state.CurrentParagraphBorderSide = null;
            _currentRowBorderSide = null;
            _pendingCellProperties.CurrentBorderSide = null;
            EnsureCurrentSection().PageSetup.PageBorders.GetBorder(side).Style = RtfPageBorderStyle.Single;
        }

        private RtfPageBorder? GetCurrentPageBorder(CharacterState state) {
            return state.CurrentPageBorderSide.HasValue
                ? EnsureCurrentSection().PageSetup.PageBorders.GetBorder(state.CurrentPageBorderSide.Value)
                : null;
        }

        private bool TryApplyCurrentPageBorderStyle(CharacterState state, RtfPageBorderStyle style, bool shadow = false) {
            RtfPageBorder? border = GetCurrentPageBorder(state);
            if (border == null) return false;
            border.Style = style;
            if (shadow) {
                border.Shadow = true;
            }

            return true;
        }

        private bool TryApplyCurrentPageBorderWidth(CharacterState state, int? width) {
            RtfPageBorder? border = GetCurrentPageBorder(state);
            if (border == null) return false;
            border.Width = width;
            return true;
        }

        private bool TryApplyCurrentPageBorderSpace(CharacterState state, int? space) {
            RtfPageBorder? border = GetCurrentPageBorder(state);
            if (border == null) return false;
            border.Space = space;
            return true;
        }

        private bool TryApplyCurrentPageBorderColor(CharacterState state, int? colorIndex) {
            RtfPageBorder? border = GetCurrentPageBorder(state);
            if (border == null) return false;
            border.ColorIndex = colorIndex;
            return true;
        }

        private bool TryApplyCurrentPageBorderFrame(CharacterState state, bool frame) {
            RtfPageBorder? border = GetCurrentPageBorder(state);
            if (border == null) return false;
            border.Frame = frame;
            return true;
        }

        private void AddDocumentBlock(IRtfBlock block) {
            CountSemanticBlock();
            _document.AddParsedBlock(block);
            EnsureCurrentSection().AddParsedBlock(block);
        }

        private void CountSemanticBlock() => _limits.AddSemanticBlock(-1);

        private void AddPicture(RtfImage image) {
            if (_currentParagraph.Inlines.Count > 0 ||
                _currentParagraphIsInTable ||
                _currentRow != null ||
                _currentNote != null ||
                _currentHeaderFooter != null ||
                _inlineCaptureDepth > 0) {
                _currentParagraph.AddImage(image);
                return;
            }

            AddDocumentBlock(image);
        }

        private void CompleteCurrentSection() {
            RtfSection section = EnsureCurrentSection();
            if (section.Blocks.Count > 0 || section.HasAnyLayoutValue) {
                _document.AddParsedSection(section);
            }

            _currentSection = new RtfSection();
            _currentSectionColumnNumber = null;
        }

        private void CompleteOpenSection() {
            if (!_hasSemanticSections || _currentSection == null) {
                return;
            }

            if (_currentSection.Blocks.Count > 0 || _currentSection.HasAnyLayoutValue) {
                _document.AddParsedSection(_currentSection);
            }
        }
    }
}
