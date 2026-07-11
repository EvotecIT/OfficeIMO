namespace OfficeIMO.Markup;

public sealed partial class OfficeMarkupCSharpEmitter {
    private static void EmitPresentation(OfficeMarkupDocument document, OfficeMarkupEmitterOptions options, StringBuilder sb) {
        sb.AppendLine("using OfficeIMO.PowerPoint;");
        sb.AppendLine("using OfficeIMO.Drawing;");
        sb.AppendLine("using C = DocumentFormat.OpenXml.Drawing.Charts;");
        sb.AppendLine();
        sb.AppendLine($"using PowerPointPresentation presentation = PowerPointPresentation.Create({options.FilePathVariable});");
        var slideIndex = 0;
        var chartIndex = 0;
        string? activeSection = null;
        foreach (var slide in GetPresentationSlides(document)) {
            slideIndex++;
            sb.AppendLine($"PowerPointSlide slide{slideIndex} = presentation.AddSlide();");
            if (!string.IsNullOrWhiteSpace(slide.Section)) {
                var section = slide.Section!.Trim();
                sb.AppendLine($"// slide{slideIndex}: section {CsString(section)}");
                if (!string.Equals(activeSection, section, StringComparison.Ordinal)) {
                    sb.AppendLine($"presentation.AddSection({CsString(section)}, startSlideIndex: {slideIndex - 1});");
                    activeSection = section;
                }
            }

            if (!string.IsNullOrWhiteSpace(slide.Transition)) {
                var resolvedTransition = OfficeMarkupTransitionResolver.Parse(slide.Transition);
                if (!string.IsNullOrWhiteSpace(resolvedTransition.ResolvedIdentifier)) {
                    sb.AppendLine($"slide{slideIndex}.Transition = SlideTransition.{resolvedTransition.ResolvedIdentifier};");
                    EmitTransitionAssignments(sb, $"slide{slideIndex}", resolvedTransition);
                }

                if (resolvedTransition.HasArguments || string.IsNullOrWhiteSpace(resolvedTransition.ResolvedIdentifier)) {
                    EmitTransitionComments(sb, resolvedTransition);
                }
            }

            if (!string.IsNullOrWhiteSpace(slide.Background)) {
                sb.AppendLine($"// slide{slideIndex}: background {CsString(slide.Background!)}");
            }

            if (!string.IsNullOrWhiteSpace(slide.Layout)) {
                sb.AppendLine($"// slide{slideIndex}: layout {CsString(slide.Layout!)}");
            }

            if (!string.IsNullOrWhiteSpace(slide.Title)) {
                sb.AppendLine($"slide{slideIndex}.AddTextBox({CsString(slide.Title!)});");
            }

            EmitContentBlocksForSlide(slide.Blocks, $"slide{slideIndex}", sb, ref chartIndex);
            if (!string.IsNullOrWhiteSpace(slide.Notes)) {
                sb.AppendLine($"slide{slideIndex}.Notes.Text = {CsString(slide.Notes!)};");
            }

            sb.AppendLine();
        }

        sb.AppendLine("presentation.Save();");
    }

    private static IEnumerable<OfficeMarkupSlideBlock> GetPresentationSlides(OfficeMarkupDocument document) {
        var pendingBlocks = new List<OfficeMarkupBlock>();
        foreach (var block in document.Blocks) {
            if (block is OfficeMarkupSlideBlock slide) {
                if (pendingBlocks.Count > 0) {
                    yield return CreateImplicitSlide(pendingBlocks);
                    pendingBlocks.Clear();
                }

                yield return slide;
            } else {
                pendingBlocks.Add(block);
            }
        }

        if (pendingBlocks.Count > 0) {
            yield return CreateImplicitSlide(pendingBlocks);
        }
    }

    private static OfficeMarkupSlideBlock CreateImplicitSlide(IReadOnlyList<OfficeMarkupBlock> blocks) {
        var slide = new OfficeMarkupSlideBlock();
        var startIndex = 0;
        if (blocks.Count > 0 && blocks[0] is OfficeMarkupHeadingBlock heading && heading.Level == 1) {
            slide.Title = heading.Text;
            startIndex = 1;
        }

        for (var index = startIndex; index < blocks.Count; index++) {
            slide.Blocks.Add(blocks[index]);
        }

        return slide;
    }

    private static void EmitContentBlocksForSlide(IEnumerable<OfficeMarkupBlock> blocks, string slideVariable, StringBuilder sb, ref int chartIndex) {
        foreach (var block in blocks) {
            switch (block) {
                case OfficeMarkupHeadingBlock heading:
                    sb.AppendLine($"{slideVariable}.AddTextBox({CsString(heading.Text)});");
                    break;
                case OfficeMarkupParagraphBlock paragraph:
                    sb.AppendLine($"{slideVariable}.AddTextBox({CsString(paragraph.Text)});");
                    break;
                case OfficeMarkupListBlock list:
                    sb.AppendLine($"{slideVariable}.AddTextBox({CsString(FormatList(list))});");
                    break;
                case OfficeMarkupImageBlock image:
                    EmitPlacementComment(image.Placement, sb);
                    sb.AppendLine($"{slideVariable}.AddPicture({CsString(image.Source)});");
                    break;
                case OfficeMarkupTableBlock table:
                    var rows = Math.Max(1, table.Rows.Count + (table.Headers.Count > 0 ? 1 : 0));
                    var columns = Math.Max(1, table.Headers.Count > 0 ? table.Headers.Count : table.Rows.Select(r => r.Count).DefaultIfEmpty(1).Max());
                    sb.AppendLine($"var table = {slideVariable}.AddTable({rows}, {columns});");
                    sb.AppendLine("// Fill table cells from the AST before saving.");
                    break;
                case OfficeMarkupDiagramBlock diagram:
                    EmitPlacementComment(diagram.Placement, sb);
                    sb.AppendLine($"// Render {diagram.Language} diagram to an image, then call {slideVariable}.AddPicture(...).");
                    break;
                case OfficeMarkupChartBlock chart:
                    EmitPowerPointChart(chart, slideVariable, sb, ++chartIndex);
                    break;
                case OfficeMarkupTextBoxBlock textBox:
                    EmitPlacementComment(textBox.Placement, sb);
                    if (!string.IsNullOrWhiteSpace(textBox.Style)) {
                        sb.AppendLine($"// Textbox style: {CsString(textBox.Style!)}");
                    }

                    sb.AppendLine($"{slideVariable}.AddTextBox({CsString(textBox.Text)});");
                    break;
                case OfficeMarkupCardBlock card:
                    EmitPlacementComment(card.Placement, sb);
                    if (!string.IsNullOrWhiteSpace(card.Style)) {
                        sb.AppendLine($"// Card style: {CsString(card.Style!)}");
                    }

                    sb.AppendLine($"{slideVariable}.AddTextBox({CsString((card.Title ?? string.Empty) + Environment.NewLine + card.Body)});");
                    break;
                case OfficeMarkupColumnsBlock columnsBlock:
                    EmitPlacementComment(columnsBlock.Placement, sb);
                    sb.AppendLine($"// Start a semantic columns region; gap={CsString(columnsBlock.Gap ?? string.Empty)}. Place following Column blocks into separate slide regions.");
                    break;
                case OfficeMarkupColumnBlock column:
                    EmitComment(sb, $"Column {column.ColumnKind} width={column.Width ?? string.Empty}");
                    if (!string.IsNullOrWhiteSpace(column.Body)) {
                        EmitComment(sb, column.Body);
                    }

                    break;
                default:
                    sb.AppendLine($"// {block.Kind}: {CsString(Describe(block))}");
                    break;
            }
        }
    }

    private static void EmitTransitionComments(StringBuilder sb, OfficeMarkupResolvedTransition resolvedTransition) {
        sb.AppendLine($"// Transition details: {CsString(resolvedTransition.RawText ?? string.Empty)}");

        if (!string.IsNullOrWhiteSpace(resolvedTransition.Effect)) {
            sb.AppendLine($"// Transition effect: {CsString(resolvedTransition.Effect!)}");
        }

        if (!string.IsNullOrWhiteSpace(resolvedTransition.ResolvedIdentifier)) {
            sb.AppendLine($"// Transition native enum: {CsString(resolvedTransition.ResolvedIdentifier!)}");
        }

        var direction = GetTransitionAttribute(resolvedTransition, "direction", "dir", "orientation", "axis", "mode");
        if (!string.IsNullOrWhiteSpace(direction)) {
            sb.AppendLine($"// Transition direction: {CsString(direction!)}");
        }

        var duration = GetTransitionAttribute(resolvedTransition, "duration");
        if (!string.IsNullOrWhiteSpace(duration)) {
            sb.AppendLine($"// Transition duration: {CsString(duration!)}");
        }

        var speed = GetTransitionAttribute(resolvedTransition, "speed", "spd");
        if (!string.IsNullOrWhiteSpace(speed)) {
            sb.AppendLine($"// Transition speed: {CsString(speed!)}");
        }

        var advanceOnClick = GetTransitionAttribute(resolvedTransition, "advance-on-click", "advanceonclick", "advance-click", "onclick", "click");
        if (!string.IsNullOrWhiteSpace(advanceOnClick)) {
            sb.AppendLine($"// Transition advance-on-click: {CsString(advanceOnClick!)}");
        }

        var advanceAfter = GetTransitionAttribute(resolvedTransition, "advance-after", "advanceafter", "after", "delay");
        if (!string.IsNullOrWhiteSpace(advanceAfter)) {
            sb.AppendLine($"// Transition advance-after: {CsString(advanceAfter!)}");
        }
    }

    private static void EmitTransitionAssignments(StringBuilder sb, string slideVariable, OfficeMarkupResolvedTransition resolvedTransition) {
        if (TryGetTransitionSpeed(resolvedTransition, out var speedIdentifier)) {
            sb.AppendLine($"{slideVariable}.TransitionSpeed = SlideTransitionSpeed.{speedIdentifier};");
        }

        if (TryGetTransitionSeconds(resolvedTransition, out var durationSeconds, "duration", "dur")) {
            sb.AppendLine($"{slideVariable}.TransitionDurationSeconds = {FormatDoubleLiteral(durationSeconds)};");
        }

        if (TryGetTransitionBoolean(resolvedTransition, out var advanceOnClick, "advance-on-click", "advanceonclick", "advance-click", "onclick", "click")) {
            sb.AppendLine($"{slideVariable}.TransitionAdvanceOnClick = {(advanceOnClick ? "true" : "false")};");
        }

        if (TryGetTransitionSeconds(resolvedTransition, out var advanceAfterSeconds, "advance-after", "advanceafter", "after", "delay")) {
            sb.AppendLine($"{slideVariable}.TransitionAdvanceAfterSeconds = {FormatDoubleLiteral(advanceAfterSeconds)};");
        }
    }

    private static string? GetTransitionAttribute(OfficeMarkupResolvedTransition resolvedTransition, params string[] names) {
        foreach (var name in names) {
            if (resolvedTransition.Attributes.TryGetValue(name, out var value) && !string.IsNullOrWhiteSpace(value)) {
                return value.Trim();
            }
        }

        return null;
    }

    private static bool TryGetTransitionSpeed(OfficeMarkupResolvedTransition resolvedTransition, out string identifier) {
        identifier = string.Empty;
        var value = GetTransitionAttribute(resolvedTransition, "speed", "spd");
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        switch (NormalizeTransitionToken(value)) {
            case "slow":
                identifier = "Slow";
                return true;
            case "medium":
            case "med":
                identifier = "Medium";
                return true;
            case "fast":
                identifier = "Fast";
                return true;
            default:
                return false;
        }
    }

    private static bool TryGetTransitionSeconds(OfficeMarkupResolvedTransition resolvedTransition, out double seconds, params string[] names) {
        seconds = default;
        var value = GetTransitionAttribute(resolvedTransition, names);
        return !string.IsNullOrWhiteSpace(value)
               && double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out seconds);
    }

    private static bool TryGetTransitionBoolean(OfficeMarkupResolvedTransition resolvedTransition, out bool enabled, params string[] names) {
        enabled = default;
        var value = GetTransitionAttribute(resolvedTransition, names);
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        switch (NormalizeTransitionToken(value)) {
            case "true":
            case "yes":
            case "on":
            case "1":
                enabled = true;
                return true;
            case "false":
            case "no":
            case "off":
            case "0":
                enabled = false;
                return true;
            default:
                return false;
        }
    }

    private static string NormalizeTransitionToken(string? value) =>
        new string((value ?? string.Empty).Where(char.IsLetterOrDigit).Select(char.ToLowerInvariant).ToArray());

    private static string FormatDoubleLiteral(double value) =>
        value.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture);
}
