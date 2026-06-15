using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

internal static partial class RtfSemanticReader {
    private sealed partial class Binder {
        private bool TryAppendField(RtfGroup group, CharacterState state, int depth) {
            RtfGroup? instruction = group.Children.OfType<RtfGroup>().FirstOrDefault(child => child.Destination == "fldinst");
            RtfGroup? result = group.Children.OfType<RtfGroup>().FirstOrDefault(child => child.Destination == "fldrslt");
            if (instruction == null || result == null) {
                return false;
            }

            string fieldInstruction = CollectPlainText(instruction, state.AnsiCodePage, state.UnicodeSkipCount).Trim();
            RtfParagraph resultParagraph = ReadInlineParagraph(result, state, depth);
            Uri? hyperlink = ParseHyperlinkInstruction(fieldInstruction);
            if (hyperlink != null) {
                AppendHyperlinkResult(resultParagraph, hyperlink);
                return true;
            }

            ApplyParagraphState(_currentParagraph, state);
            var field = new RtfField(fieldInstruction);
            RtfGroup? formFieldDataGroup = group.Children.OfType<RtfGroup>().FirstOrDefault(child => child.Destination == "ffdata");
            if (formFieldDataGroup != null) {
                field.FormFieldData = ReadFormFieldData(formFieldDataGroup, state);
            }

            foreach (IRtfInline inline in resultParagraph.Inlines) {
                field.Result.AddInline(inline);
            }

            _currentParagraph.AddField(field);
            return true;
        }

        private static RtfFormFieldData ReadFormFieldData(RtfGroup group, CharacterState state) {
            var formFieldData = new RtfFormFieldData();
            foreach (RtfNode node in group.Children) {
                switch (node) {
                    case RtfControlWord control when IsFormFieldControl(control.Name):
                        formFieldData.AddParsedControl(new RtfFormFieldDataControl(control.Name, control.Parameter, control.HasParameter));
                        break;
                    case RtfGroup childGroup:
                        ReadFormFieldTextDestination(childGroup, formFieldData, state);
                        break;
                }
            }

            return formFieldData;
        }

        private static bool IsFormFieldControl(string name) => name.StartsWith("ff", StringComparison.Ordinal) && name != "ffdata";

        private static void ReadFormFieldTextDestination(RtfGroup group, RtfFormFieldData formFieldData, CharacterState state) {
            string value = CollectPlainText(group, state.AnsiCodePage, state.UnicodeSkipCount);
            switch (group.Destination) {
                case "ffname":
                    formFieldData.Name = value;
                    break;
                case "ffdeftext":
                    formFieldData.DefaultText = value;
                    break;
                case "ffformat":
                    formFieldData.Format = value;
                    break;
                case "ffhelptext":
                    formFieldData.HelpText = value;
                    break;
                case "ffstattext":
                    formFieldData.StatusText = value;
                    break;
                case "ffentrymcr":
                    formFieldData.EntryMacro = value;
                    break;
                case "ffexitmcr":
                    formFieldData.ExitMacro = value;
                    break;
                case "ffl":
                    formFieldData.AddDropDownItem(value);
                    break;
            }
        }

        private RtfParagraph ReadInlineParagraph(RtfGroup group, CharacterState state, int depth) {
            RtfParagraph savedParagraph = _currentParagraph;
            RtfTable? savedTable = _currentTable;
            RtfTableRow? savedRow = _currentRow;
            RtfHeaderFooter? savedHeaderFooter = _currentHeaderFooter;
            RtfNote? savedNote = _currentNote;
            int savedCellIndex = _currentCellIndex;
            bool savedParagraphIsInTable = _currentParagraphIsInTable;

            _currentParagraph = new RtfParagraph();
            _currentTable = null;
            _currentRow = null;
            _currentHeaderFooter = null;
            _currentNote = null;
            _currentCellIndex = 0;
            _currentParagraphIsInTable = false;
            _inlineCaptureDepth++;

            try {
                WalkGroup(group, state.Clone(), depth + 1, allowDestinationSkip: false);
                return _currentParagraph;
            } finally {
                _inlineCaptureDepth--;
                _currentParagraph = savedParagraph;
                _currentTable = savedTable;
                _currentRow = savedRow;
                _currentHeaderFooter = savedHeaderFooter;
                _currentNote = savedNote;
                _currentCellIndex = savedCellIndex;
                _currentParagraphIsInTable = savedParagraphIsInTable;
            }
        }

        private void AppendHyperlinkResult(RtfParagraph resultParagraph, Uri hyperlink) {
            RtfRun? previousRun = null;
            foreach (IRtfInline inline in resultParagraph.Inlines) {
                if (inline is RtfRun run) {
                    run.Hyperlink = hyperlink;
                    if (previousRun != null && CanMergeRuns(previousRun, run)) {
                        previousRun.Text += run.Text;
                        continue;
                    }

                    _currentParagraph.AddRun(run);
                    previousRun = run;
                    continue;
                }

                _currentParagraph.AddInline(inline);
                previousRun = null;
            }
        }

        private static bool CanMergeRuns(RtfRun left, RtfRun right) {
            return left.Bold == right.Bold &&
                   left.Italic == right.Italic &&
                   left.UnderlineStyle == right.UnderlineStyle &&
                   left.Strike == right.Strike &&
                   left.DoubleStrike == right.DoubleStrike &&
                   left.Hidden == right.Hidden &&
                   left.Outline == right.Outline &&
                   left.Shadow == right.Shadow &&
                   left.Emboss == right.Emboss &&
                   left.Imprint == right.Imprint &&
                   left.CapsStyle == right.CapsStyle &&
                   left.VerticalPosition == right.VerticalPosition &&
                   Nullable.Equals(left.FontSize, right.FontSize) &&
                   left.FontId == right.FontId &&
                   left.ForegroundColorIndex == right.ForegroundColorIndex &&
                   left.HighlightColorIndex == right.HighlightColorIndex &&
                   left.UnderlineColorIndex == right.UnderlineColorIndex &&
                   left.CharacterSpacingTwips == right.CharacterSpacingTwips &&
                   left.CharacterScalePercent == right.CharacterScalePercent &&
                   left.KerningHalfPoints == right.KerningHalfPoints &&
                   left.CharacterOffsetHalfPoints == right.CharacterOffsetHalfPoints &&
                   left.StyleId == right.StyleId &&
                   Equals(left.Hyperlink, right.Hyperlink) &&
                   ReferenceEquals(left.Note, right.Note);
        }

        private static Uri? ParseHyperlinkInstruction(string instruction) {
            if (!instruction.StartsWith("HYPERLINK", StringComparison.OrdinalIgnoreCase)) {
                return null;
            }

            int firstQuote = instruction.IndexOf('"');
            int lastQuote = instruction.LastIndexOf('"');
            if (firstQuote >= 0 && lastQuote > firstQuote) {
                string target = instruction.Substring(firstQuote + 1, lastQuote - firstQuote - 1);
                return Uri.TryCreate(target, UriKind.RelativeOrAbsolute, out Uri? uri) ? uri : null;
            }

            string[] parts = instruction.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 2) return null;
            return Uri.TryCreate(parts[1], UriKind.RelativeOrAbsolute, out Uri? fallback) ? fallback : null;
        }
    }
}
