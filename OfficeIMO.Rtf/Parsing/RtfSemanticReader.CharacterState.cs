namespace OfficeIMO.Rtf;

internal static partial class RtfSemanticReader {
    private sealed class CharacterState {
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public RtfUnderlineStyle UnderlineStyle { get; set; } = RtfUnderlineStyle.None;
        public bool Strike { get; set; }
        public bool DoubleStrike { get; set; }
        public bool Hidden { get; set; }
        public bool Outline { get; set; }
        public bool Shadow { get; set; }
        public bool Emboss { get; set; }
        public bool Imprint { get; set; }
        public RtfCapsStyle CapsStyle { get; set; } = RtfCapsStyle.None;
        public RtfVerticalPosition VerticalPosition { get; set; } = RtfVerticalPosition.Baseline;
        public double? FontSize { get; set; }
        public int? FontId { get; set; }
        public int? ForegroundColorIndex { get; set; }
        public int? HighlightColorIndex { get; set; }
        public int? CharacterBackgroundColorIndex { get; set; }
        public int? CharacterShadingForegroundColorIndex { get; set; }
        public int? CharacterShadingPatternPercent { get; set; }
        public RtfShadingPattern CharacterShadingPattern { get; set; } = RtfShadingPattern.None;
        public RtfCharacterBorder CharacterBorder { get; } = new RtfCharacterBorder();
        public bool CurrentCharacterBorderActive { get; set; }
        public int? UnderlineColorIndex { get; set; }
        public int? CharacterSpacingTwips { get; set; }
        public int? CharacterScalePercent { get; set; }
        public int? KerningHalfPoints { get; set; }
        public int? CharacterOffsetHalfPoints { get; set; }
        public int? CharacterStyleId { get; set; }
        public RtfTextDirection? Direction { get; set; }
        public int? DefaultLanguageId { get; set; }
        public int? LanguageId { get; set; }
        public RtfRevisionKind RevisionKind { get; set; } = RtfRevisionKind.None;
        public int? RevisionAuthorIndex { get; set; }
        public int? RevisionTimestampValue { get; set; }
        public int? CharacterRevisionSaveId { get; set; }
        public int? InsertionRevisionSaveId { get; set; }
        public int? DeletionRevisionSaveId { get; set; }
        public RtfTextAlignment Alignment { get; set; } = RtfTextAlignment.Left;
        public RtfTextDirection? ParagraphDirection { get; set; }
        public int? ParagraphStyleId { get; set; }
        public int? ListId { get; set; }
        public int? ListDefinitionId { get; set; }
        public int? ListLevel { get; set; }
        public RtfListKind ListKind { get; set; } = RtfListKind.None;
        public RtfLegacyNumbering LegacyNumbering { get; } = new RtfLegacyNumbering();
        public RtfLegacyNumbering PendingLegacyNumberingAfterReset { get; } = new RtfLegacyNumbering();
        public bool HasPendingLegacyNumberingAfterReset { get; set; }
        public RtfParagraph? ListText { get; set; }
        public RtfParagraph? PendingListTextAfterReset { get; set; }
        public int? LeftIndentTwips { get; set; }
        public int? RightIndentTwips { get; set; }
        public int? FirstLineIndentTwips { get; set; }
        public int? SpaceBeforeTwips { get; set; }
        public int? SpaceAfterTwips { get; set; }
        public bool? SpaceBeforeAuto { get; set; }
        public bool? SpaceAfterAuto { get; set; }
        public int? LineSpacingTwips { get; set; }
        public bool? LineSpacingMultiple { get; set; }
        public int? BackgroundColorIndex { get; set; }
        public int? ShadingForegroundColorIndex { get; set; }
        public int? ShadingPatternPercent { get; set; }
        public RtfShadingPattern ShadingPattern { get; set; } = RtfShadingPattern.None;
        public RtfParagraphBorderSide? CurrentParagraphBorderSide { get; set; }

        public RtfPageBorderSide? CurrentPageBorderSide { get; set; }
        public RtfParagraphBorder TopBorder { get; } = new RtfParagraphBorder();
        public RtfParagraphBorder LeftBorder { get; } = new RtfParagraphBorder();
        public RtfParagraphBorder BottomBorder { get; } = new RtfParagraphBorder();
        public RtfParagraphBorder RightBorder { get; } = new RtfParagraphBorder();
        public bool PageBreakBefore { get; set; }
        public bool KeepWithNext { get; set; }
        public bool KeepLinesTogether { get; set; }
        public bool SuppressLineNumbers { get; set; }
        public bool? AutoHyphenation { get; set; }
        public bool? ContextualSpacing { get; set; }
        public bool? AdjustRightIndent { get; set; }
        public bool? SnapToLineGrid { get; set; }
        public bool? WidowControl { get; set; }
        public int? OutlineLevel { get; set; }
        public int? ParagraphRevisionSaveId { get; set; }
        public RtfParagraphFrame Frame { get; } = new RtfParagraphFrame();
        public List<RtfTabStop> TabStops { get; } = new List<RtfTabStop>();
        public RtfTabAlignment PendingTabAlignment { get; set; } = RtfTabAlignment.Left;
        public RtfTabLeader PendingTabLeader { get; set; } = RtfTabLeader.None;
        public int AnsiCodePage { get; set; } = RtfAnsiCodePage.DefaultWindowsCodePage;
        public bool HasExplicitAnsiCodePage { get; set; }
        public int UnicodeSkipCount { get; set; } = 1;
        public int SkipCharacters { get; set; }
        public char? PendingHighSurrogate { get; set; }
        public byte? PendingAnsiLeadByte { get; set; }

        public CharacterState Clone() {
            var clone = new CharacterState {
                Bold = Bold,
                Italic = Italic,
                UnderlineStyle = UnderlineStyle,
                Strike = Strike,
                DoubleStrike = DoubleStrike,
                Hidden = Hidden,
                Outline = Outline,
                Shadow = Shadow,
                Emboss = Emboss,
                Imprint = Imprint,
                CapsStyle = CapsStyle,
                VerticalPosition = VerticalPosition,
                FontSize = FontSize,
                FontId = FontId,
                ForegroundColorIndex = ForegroundColorIndex,
                HighlightColorIndex = HighlightColorIndex,
                CharacterBackgroundColorIndex = CharacterBackgroundColorIndex,
                CharacterShadingForegroundColorIndex = CharacterShadingForegroundColorIndex,
                CharacterShadingPatternPercent = CharacterShadingPatternPercent,
                CharacterShadingPattern = CharacterShadingPattern,
                CurrentCharacterBorderActive = CurrentCharacterBorderActive,
                UnderlineColorIndex = UnderlineColorIndex,
                CharacterSpacingTwips = CharacterSpacingTwips,
                CharacterScalePercent = CharacterScalePercent,
                KerningHalfPoints = KerningHalfPoints,
                CharacterOffsetHalfPoints = CharacterOffsetHalfPoints,
                CharacterStyleId = CharacterStyleId,
                Direction = Direction,
                DefaultLanguageId = DefaultLanguageId,
                LanguageId = LanguageId,
                RevisionKind = RevisionKind,
                RevisionAuthorIndex = RevisionAuthorIndex,
                RevisionTimestampValue = RevisionTimestampValue,
                CharacterRevisionSaveId = CharacterRevisionSaveId,
                InsertionRevisionSaveId = InsertionRevisionSaveId,
                DeletionRevisionSaveId = DeletionRevisionSaveId,
                Alignment = Alignment,
                ParagraphDirection = ParagraphDirection,
                ParagraphStyleId = ParagraphStyleId,
                ListId = ListId,
                ListDefinitionId = ListDefinitionId,
                ListLevel = ListLevel,
                ListKind = ListKind,
                HasPendingLegacyNumberingAfterReset = HasPendingLegacyNumberingAfterReset,
                ListText = ListText,
                PendingListTextAfterReset = PendingListTextAfterReset,
                LeftIndentTwips = LeftIndentTwips,
                RightIndentTwips = RightIndentTwips,
                FirstLineIndentTwips = FirstLineIndentTwips,
                SpaceBeforeTwips = SpaceBeforeTwips,
                SpaceAfterTwips = SpaceAfterTwips,
                SpaceBeforeAuto = SpaceBeforeAuto,
                SpaceAfterAuto = SpaceAfterAuto,
                LineSpacingTwips = LineSpacingTwips,
                LineSpacingMultiple = LineSpacingMultiple,
                BackgroundColorIndex = BackgroundColorIndex,
                ShadingForegroundColorIndex = ShadingForegroundColorIndex,
                ShadingPatternPercent = ShadingPatternPercent,
                ShadingPattern = ShadingPattern,
                CurrentParagraphBorderSide = CurrentParagraphBorderSide,
                CurrentPageBorderSide = CurrentPageBorderSide,
                PageBreakBefore = PageBreakBefore,
                KeepWithNext = KeepWithNext,
                KeepLinesTogether = KeepLinesTogether,
                SuppressLineNumbers = SuppressLineNumbers,
                AutoHyphenation = AutoHyphenation,
                ContextualSpacing = ContextualSpacing,
                AdjustRightIndent = AdjustRightIndent,
                SnapToLineGrid = SnapToLineGrid,
                WidowControl = WidowControl,
                OutlineLevel = OutlineLevel,
                ParagraphRevisionSaveId = ParagraphRevisionSaveId,
                PendingTabAlignment = PendingTabAlignment,
                PendingTabLeader = PendingTabLeader,
                AnsiCodePage = AnsiCodePage,
                HasExplicitAnsiCodePage = HasExplicitAnsiCodePage,
                UnicodeSkipCount = UnicodeSkipCount,
                SkipCharacters = SkipCharacters,
                PendingHighSurrogate = PendingHighSurrogate,
                PendingAnsiLeadByte = PendingAnsiLeadByte
            };
            clone.CharacterBorder.CopyFrom(CharacterBorder);
            clone.LegacyNumbering.CopyFrom(LegacyNumbering);
            clone.PendingLegacyNumberingAfterReset.CopyFrom(PendingLegacyNumberingAfterReset);
            clone.Frame.CopyFrom(Frame);
            foreach (RtfTabStop tabStop in TabStops) {
                clone.TabStops.Add(new RtfTabStop(tabStop.PositionTwips, tabStop.Alignment, tabStop.Leader));
            }

            CopyBorder(TopBorder, clone.TopBorder);
            CopyBorder(LeftBorder, clone.LeftBorder);
            CopyBorder(BottomBorder, clone.BottomBorder);
            CopyBorder(RightBorder, clone.RightBorder);
            return clone;
        }

        public void ResetCharacter() {
            Bold = false;
            Italic = false;
            UnderlineStyle = RtfUnderlineStyle.None;
            Strike = false;
            DoubleStrike = false;
            Hidden = false;
            Outline = false;
            Shadow = false;
            Emboss = false;
            Imprint = false;
            CapsStyle = RtfCapsStyle.None;
            VerticalPosition = RtfVerticalPosition.Baseline;
            FontSize = null;
            FontId = null;
            ForegroundColorIndex = null;
            HighlightColorIndex = null;
            CharacterBackgroundColorIndex = null;
            CharacterShadingForegroundColorIndex = null;
            CharacterShadingPatternPercent = null;
            CharacterShadingPattern = RtfShadingPattern.None;
            CharacterBorder.Clear();
            CurrentCharacterBorderActive = false;
            UnderlineColorIndex = null;
            CharacterSpacingTwips = null;
            CharacterScalePercent = null;
            KerningHalfPoints = null;
            CharacterOffsetHalfPoints = null;
            CharacterStyleId = null;
            Direction = null;
            LanguageId = DefaultLanguageId;
            RevisionKind = RtfRevisionKind.None;
            RevisionAuthorIndex = null;
            RevisionTimestampValue = null;
            CharacterRevisionSaveId = null;
            InsertionRevisionSaveId = null;
            DeletionRevisionSaveId = null;
        }

        public void ResetParagraph() {
            ResetCharacter();
            Alignment = RtfTextAlignment.Left;
            ParagraphDirection = null;
            ParagraphStyleId = null;
            ListId = null;
            ListDefinitionId = null;
            ListLevel = null;
            ListKind = RtfListKind.None;
            LegacyNumbering.Clear();
            PendingLegacyNumberingAfterReset.Clear();
            HasPendingLegacyNumberingAfterReset = false;
            ListText = null;
            PendingListTextAfterReset = null;
            LeftIndentTwips = null;
            RightIndentTwips = null;
            FirstLineIndentTwips = null;
            SpaceBeforeTwips = null;
            SpaceAfterTwips = null;
            SpaceBeforeAuto = null;
            SpaceAfterAuto = null;
            LineSpacingTwips = null;
            LineSpacingMultiple = null;
            BackgroundColorIndex = null;
            ShadingForegroundColorIndex = null;
            ShadingPatternPercent = null;
            ShadingPattern = RtfShadingPattern.None;
            CurrentParagraphBorderSide = null;
            CurrentPageBorderSide = null;
            ClearBorder(TopBorder);
            ClearBorder(LeftBorder);
            ClearBorder(BottomBorder);
            ClearBorder(RightBorder);
            PageBreakBefore = false;
            KeepWithNext = false;
            KeepLinesTogether = false;
            SuppressLineNumbers = false;
            AutoHyphenation = null;
            ContextualSpacing = null;
            AdjustRightIndent = null;
            SnapToLineGrid = null;
            WidowControl = null;
            OutlineLevel = null;
            ParagraphRevisionSaveId = null;
            Frame.Clear();
            TabStops.Clear();
            PendingTabAlignment = RtfTabAlignment.Left;
            PendingTabLeader = RtfTabLeader.None;
        }

        public void AddTabStop(int positionTwips, RtfTabAlignment alignment) {
            TabStops.Add(new RtfTabStop(positionTwips, alignment, PendingTabLeader));
            PendingTabAlignment = RtfTabAlignment.Left;
            PendingTabLeader = RtfTabLeader.None;
        }

        public RtfParagraphBorder GetParagraphBorder(RtfParagraphBorderSide side) {
            switch (side) {
                case RtfParagraphBorderSide.Top:
                    return TopBorder;
                case RtfParagraphBorderSide.Left:
                    return LeftBorder;
                case RtfParagraphBorderSide.Bottom:
                    return BottomBorder;
                default:
                    return RightBorder;
            }
        }

        public RtfParagraphBorder? GetCurrentParagraphBorder() {
            return CurrentParagraphBorderSide.HasValue ? GetParagraphBorder(CurrentParagraphBorderSide.Value) : null;
        }

        private static void CopyBorder(RtfParagraphBorder source, RtfParagraphBorder destination) {
            destination.Style = source.Style;
            destination.Width = source.Width;
            destination.ColorIndex = source.ColorIndex;
        }

        private static void ClearBorder(RtfParagraphBorder border) {
            border.Style = RtfParagraphBorderStyle.None;
            border.Width = null;
            border.ColorIndex = null;
        }
    }
}
