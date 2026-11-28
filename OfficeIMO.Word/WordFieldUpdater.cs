using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word.OpenXml;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Word {
    internal static partial class WordFieldUpdater {
        private const int FieldNamespace = OfficeXmlNamespaces.WordprocessingML;

        internal static WordFieldUpdateReport Update(WordDocument document, WordFieldUpdateOptions options) {
            if (options.DateTimeFieldsOnly && options.TablesOfContentsOnly) {
                throw new InvalidOperationException("DateTimeFieldsOnly and TablesOfContentsOnly each select the only fields to refresh. Set one of them.");
            }

            if (!options.DateTimeFieldsOnly && options.UseLayoutPagination && options.LayoutPaginationProvider == null) {
                throw new InvalidOperationException("Layout pagination requires OfficeIMO.Word.Pdf. Select its engine with WithPdfLayoutPagination() on the field update options.");
            }

            bool layOut = !options.DateTimeFieldsOnly && options.UseLayoutPagination;
            bool refreshToc = !options.DateTimeFieldsOnly && (options.UpdateTablesOfContents || options.TablesOfContentsOnly);
            if (layOut) document.OpenSdkPackage("Layout field evaluation");
            var state = new FieldEvaluationState();
            List<MutableFieldCandidate> candidates = EnumerateFields(document);
            if (!options.DateTimeFieldsOnly && document._ownedModel is { } owned &&
                candidates.Any(candidate => RequiresSdkEvaluation(owned, candidate))) {
                document.OpenSdkPackage("Field evaluation needs document properties");
                candidates = EnumerateFields(document);
            }
            state.LockedTocFields.AddRange(candidates.Where(IsLockedToc));
            if (refreshToc) {
                RegenerateTablesOfContents(document, state, WordTableOfContent.GetTablesOfContents(document));
                if (state.RegeneratedTablesOfContents.Count != 0) candidates = EnumerateFields(document);
            }
            DateTime updateDateTime = options.CurrentDateTime ?? DateTime.Now;
            WordFieldEvaluationReason dateTimeEvaluationReason = options.CurrentDateTime.HasValue
                ? WordFieldEvaluationReason.CallerProvidedDateTime
                : WordFieldEvaluationReason.RuntimeClock;
            var results = new List<WordFieldUpdateResult>();
            var replacedContainingFields = new List<MutableFieldCandidate>();

            for (int index = 0; index < candidates.Count; index++) {
                candidates[index].Index = index;
            }

            IReadOnlyList<MutableFieldCandidate> ordered = SortFieldsForEvaluation(candidates);
            if (!layOut) {
                foreach (MutableFieldCandidate candidate in ordered) Process(candidate);
                return new WordFieldUpdateReport(results.OrderBy(result => result.Index).ToArray());
            }

            // Fields whose text can change the layout (DOCVARIABLE, REF, dates and so on) are updated first, so the
            // layout measures their new text. Page fields then take their numbers from that one layout. Like Word's own
            // update, there is no second pass when a new page number changes line wrapping.
            HashSet<int> pageDependent = GetPageDependentFields(candidates);
            foreach (MutableFieldCandidate candidate in ordered.Where(candidate => !pageDependent.Contains(candidate.Sequence))) Process(candidate);
            state.Layout = LayOutWithPageFieldBookmarks(document, options.LayoutPaginationProvider!, state);
            foreach (MutableFieldCandidate candidate in ordered.Where(candidate => pageDependent.Contains(candidate.Sequence))) Process(candidate);
            return new WordFieldUpdateReport(results.OrderBy(result => result.Index).ToArray());

            void Process(MutableFieldCandidate candidate) {
                if (options.DateTimeFieldsOnly) {
                    WordFieldType? fieldType = WordFieldInventory.ParseInstruction(candidate.InstructionText).FieldType;
                    if (fieldType is not (WordFieldType.Date or WordFieldType.Time)) {
                        results.Add(candidate.ToResult(fieldType, WordFieldUpdateStatus.Skipped, null,
                            "Field was left unchanged because only DATE and TIME fields were selected."));
                        return;
                    }
                }
                if (IsInsideRegeneratedTableOfContents(candidate, state)) {
                    WordFieldType? fieldType = WordFieldInventory.ParseInstruction(candidate.InstructionText).FieldType;
                    if (fieldType == WordFieldType.TOC) {
                        results.Add(candidate.ToResult(fieldType, WordFieldUpdateStatus.Updated, null,
                            "Table of contents was regenerated from the document headings; its page numbers are PAGEREF fields.",
                            WordFieldEvaluationReason.TableOfContentsRegenerated));
                        return;
                    }
                }

                if (IsNestedInsideReplacedField(candidate, replacedContainingFields)) {
                    results.Add(candidate.ToResult(
                        WordFieldInventory.ParseInstruction(candidate.InstructionText).FieldType,
                        WordFieldUpdateStatus.Skipped,
                        null,
                        "Nested field was left unchanged because its containing field result was replaced.",
                        WordFieldEvaluationReason.ContainingResultReplaced));
                    return;
                }

                if (IsInsideLockedToc(candidate, state.LockedTocFields)) {
                    WordFieldType? lockedType = WordFieldInventory.ParseInstruction(candidate.InstructionText).FieldType;
                    results.Add(candidate.ToResult(lockedType, WordFieldUpdateStatus.Skipped, null, "Field is inside a locked table of contents and was left unchanged."));
                    return;
                }

                if (options.TablesOfContentsOnly && !IsInsideRegeneratedTableOfContents(candidate, state)) {
                    results.Add(candidate.ToResult(
                        WordFieldInventory.ParseInstruction(candidate.InstructionText).FieldType,
                        WordFieldUpdateStatus.Skipped,
                        null,
                        "Field was left unchanged because only tables of contents were selected."));
                    return;
                }

                WordFieldUpdateResult result = UpdateField(
                    document, candidate, state, updateDateTime, dateTimeEvaluationReason);
                results.Add(result);
                if (result.Status == WordFieldUpdateStatus.Updated &&
                    candidate.Representation == WordFieldRepresentation.Complex &&
                    candidate.EndRun.Exists) {
                    replacedContainingFields.Add(candidate);
                }
            }
        }

        private static bool RequiresSdkEvaluation(WordPackageModel model, MutableFieldCandidate candidate) {
            // A nested field can change the command or property that the outer field needs.
            if (candidate.HasDynamicInstruction) return true;
            var parsed = WordFieldInventory.ParseInstruction(candidate.InstructionText);
            return parsed.FieldType switch {
                WordFieldType.DocProperty => CustomFieldUsesUnsupportedVariant(model, parsed),
                WordFieldType.NumChars => !ReadStoredFieldStatistic(model, "Characters").HasValue,
                WordFieldType.NumWords => !ReadStoredFieldStatistic(model, "Words").HasValue,
                _ => false
            };
        }

        /// <summary>Page fields, fields whose instruction contains one, and the sequences they belong to take their result after the layout.</summary>
        private static HashSet<int> GetPageDependentFields(IReadOnlyList<MutableFieldCandidate> candidates) {
            var bySequence = candidates.ToDictionary(candidate => candidate.Sequence);
            var dependent = new HashSet<int>();
            foreach (MutableFieldCandidate candidate in candidates) {
                if (WordFieldInventory.ParseInstruction(candidate.InstructionText).FieldType is not
                    (WordFieldType.Page or WordFieldType.NumPages or WordFieldType.PageRef or WordFieldType.SectionPages)) continue;
                AddWithOwners(candidate);
            }

            // SEQ and LISTNUM fields number from the fields before them. When one of them waits for the layout (a reset
            // value taken from a PAGE field), every field of that kind waits too, so they still run in document order.
            // Deferring one group can defer an owner from the other group, so repeat until nothing is added. Each pass
            // only adds fields, so this ends after at most one pass per group plus one.
            var groups = new[] { WordFieldType.Seq, WordFieldType.ListNum }.Select(stateful => candidates
                .Where(candidate => WordFieldInventory.ParseInstruction(candidate.InstructionText).FieldType == stateful).ToList()).ToList();
            var deferredGroups = new HashSet<int>();
            bool added = true;
            while (added) {
                added = false;
                for (int group = 0; group < groups.Count; group++) {
                    if (deferredGroups.Contains(group) || !groups[group].Any(candidate => dependent.Contains(candidate.Sequence))) continue;
                    deferredGroups.Add(group);
                    foreach (MutableFieldCandidate candidate in groups[group]) AddWithOwners(candidate);
                    added = true;
                }
            }
            return dependent;

            void AddWithOwners(MutableFieldCandidate candidate) {
                for (MutableFieldCandidate? current = candidate; current != null && dependent.Add(current.Sequence);
                     current = current.InstructionParentSequence is int parent && bySequence.TryGetValue(parent, out MutableFieldCandidate? owner) ? owner : null) {
                }
            }
        }

        internal static void UpdatePageCounters(WordDocument document) {
            var state = new FieldEvaluationState();
            List<MutableFieldCandidate> candidates = EnumerateFields(document);
            state.LockedTocFields.AddRange(candidates.Where(IsLockedToc));
            foreach (MutableFieldCandidate candidate in candidates) {
                WordFieldInventory.ParsedFieldInstruction parsed = WordFieldInventory.ParseInstruction(candidate.InstructionText);
                if (parsed.Diagnostics.Count > 0 ||
                    parsed.FieldType is not (WordFieldType.Page or WordFieldType.NumPages) ||
                    candidate.IsLocked || IsInsideLockedToc(candidate, state.LockedTocFields)) {
                    continue;
                }

                if (TryEvaluate(document, candidate, parsed, state, DateTime.Now, out string? value, out WordFieldUpdateStatus status, out _, out _) &&
                    status == WordFieldUpdateStatus.Updated) {
                    SetResultText(document, candidate, value!);
                }
            }

        }

        private static WordFieldUpdateResult UpdateField(
            WordDocument document,
            MutableFieldCandidate candidate,
            FieldEvaluationState state,
            DateTime updateDateTime,
            WordFieldEvaluationReason dateTimeEvaluationReason) {
            WordFieldInventory.ParsedFieldInstruction parsed = WordFieldInventory.ParseInstruction(candidate.InstructionText);

            if (parsed.Diagnostics.Count > 0 || parsed.FieldType == null) {
                WordFieldUpdateStatus parseStatus = parsed.FieldType == null
                    ? WordFieldUpdateStatus.ParseError
                    : WordFieldUpdateStatus.Unsupported;
                return candidate.ToResult(
                    parsed.FieldType,
                    parseStatus,
                    null,
                    parsed.Diagnostics.Count == 0 ? "Field instruction could not be parsed." : string.Join(" ", parsed.Diagnostics),
                    GetFailedEvaluationReason(candidate, parsed.FieldType, parseStatus, parsed.Diagnostics.Count == 0 ? string.Empty : string.Join(" ", parsed.Diagnostics)));
            }

            if (candidate.IsLocked) {
                return candidate.ToResult(parsed.FieldType, WordFieldUpdateStatus.Skipped, null, "Field is locked and was left unchanged.");
            }

            if (!TryEvaluate(document, candidate, parsed, state, updateDateTime, out string? value, out WordFieldUpdateStatus status, out string message, out bool usedLayout)) {
                return candidate.ToResult(parsed.FieldType, status, null, message, GetFailedEvaluationReason(candidate, parsed.FieldType, status, message));
            }

            SetResultText(document, candidate, value!);
            WordFieldEvaluationReason reason = parsed.FieldType is WordFieldType.Date or WordFieldType.Time
                ? dateTimeEvaluationReason
                : usedLayout
                ? WordFieldEvaluationReason.LayoutPagination
                : WordFieldEvaluationReason.Default;
            return candidate.ToResult(parsed.FieldType, WordFieldUpdateStatus.Updated, value, message, reason);
        }

        private static WordFieldEvaluationReason GetFailedEvaluationReason(
            MutableFieldCandidate candidate,
            WordFieldType? fieldType,
            WordFieldUpdateStatus status,
            string message) {
            if (candidate.NestingLevel > 0 &&
                fieldType == WordFieldType.ListNum &&
                message.IndexOf("Nested LISTNUM", StringComparison.OrdinalIgnoreCase) >= 0) {
                return WordFieldEvaluationReason.NestedInstruction;
            }

            if (fieldType == WordFieldType.ListNum) {
                return WordFieldEvaluationReason.ListNumberingProfileUnsupported;
            }

            if (message.IndexOf("nested table", StringComparison.OrdinalIgnoreCase) >= 0 ||
                message.IndexOf("complex table geometry", StringComparison.OrdinalIgnoreCase) >= 0) {
                return WordFieldEvaluationReason.ComplexTableProfileUnsupported;
            }

            if (message.IndexOf("locale-specific", StringComparison.OrdinalIgnoreCase) >= 0 ||
                message.IndexOf("locale-dependent", StringComparison.OrdinalIgnoreCase) >= 0) {
                return WordFieldEvaluationReason.LocaleProfileUnsupported;
            }

            if (message.IndexOf("layout", StringComparison.OrdinalIgnoreCase) >= 0 ||
                status == WordFieldUpdateStatus.Skipped && fieldType is WordFieldType.TOC or WordFieldType.Index) {
                return WordFieldEvaluationReason.ExternalLayoutRequired;
            }

            return WordFieldEvaluationReason.Default;
        }

        private static bool TryEvaluate(
            WordDocument document,
            MutableFieldCandidate candidate,
            WordFieldInventory.ParsedFieldInstruction parsed,
            FieldEvaluationState state,
            DateTime updateDateTime,
            out string? value,
            out WordFieldUpdateStatus status,
            out string message,
            out bool usedLayout) {
            usedLayout = false;
            value = null;
            status = WordFieldUpdateStatus.Unsupported;
            message = $"Field type {parsed.FieldType} is not evaluated by OfficeIMO.";

            switch (parsed.FieldType) {
                case WordFieldType.Author:
                    return TryEvaluateCoreProperty(document, CoreFieldProperty.Creator, parsed, out value, out status, out message);
                case WordFieldType.Title:
                    return TryEvaluateCoreProperty(document, CoreFieldProperty.Title, parsed, out value, out status, out message);
                case WordFieldType.Subject:
                    return TryEvaluateCoreProperty(document, CoreFieldProperty.Subject, parsed, out value, out status, out message);
                case WordFieldType.Keywords:
                    return TryEvaluateCoreProperty(document, CoreFieldProperty.Keywords, parsed, out value, out status, out message);
                case WordFieldType.Comments:
                    return TryEvaluateCoreProperty(document, CoreFieldProperty.Description, parsed, out value, out status, out message);
                case WordFieldType.LastSavedBy:
                    return TryEvaluateCoreProperty(document, CoreFieldProperty.LastModifiedBy, parsed, out value, out status, out message);
                case WordFieldType.CreateDate:
                    return TryEvaluateCoreProperty(document, CoreFieldProperty.Created, parsed, out value, out status, out message);
                case WordFieldType.Date:
                    return TrySetDate(updateDateTime, parsed, "Updated from field update date/time.", out value, out status, out message);
                case WordFieldType.PrintDate:
                    return TryEvaluateCoreProperty(document, CoreFieldProperty.LastPrinted, parsed, out value, out status, out message);
                case WordFieldType.SaveDate:
                    return TryEvaluateCoreProperty(document, CoreFieldProperty.Modified, parsed, out value, out status, out message);
                case WordFieldType.Time:
                    return TrySetDate(updateDateTime, parsed, "Updated from field update date/time.", out value, out status, out message);
                case WordFieldType.FileName:
                    return TryEvaluateFileName(document, parsed, out value, out status, out message);
                case WordFieldType.FileSize:
                    return TryEvaluateFileSize(document, parsed, out value, out status, out message);
                case WordFieldType.Info:
                    return TryEvaluateInfo(document, parsed, out value, out status, out message);
                case WordFieldType.DocProperty:
                    return TryEvaluateDocumentProperty(document, parsed, out value, out status, out message);
                case WordFieldType.DocVariable:
                    return TryEvaluateDocumentVariable(document, parsed, out value, out status, out message);
                case WordFieldType.Quote:
                    return TryEvaluateQuote(candidate.InstructionText, parsed, out value, out status, out message);
                case WordFieldType.Ref:
                    return TryEvaluateRef(document, candidate, parsed, state, out value, out status, out message);
                case WordFieldType.PageRef:
                    return TryEvaluatePageRef(document, parsed, state.Layout, out value, out status, out message, out usedLayout);
                case WordFieldType.Formula:
                    return TryEvaluateFormula(candidate, parsed, out value, out status, out message);
                case WordFieldType.Seq:
                    return TryEvaluateSequence(document, candidate, parsed, state, out value, out status, out message);
                case WordFieldType.ListNum:
                    return TryEvaluateListNum(document, candidate, parsed, state, out value, out status, out message);
                case WordFieldType.Page:
                    if (state.Layout != null && state.PageFieldBookmarks.TryGetValue(candidate.AnchorElement, out string? pageBookmark)) {
                        usedLayout = true;
                        return TryEvaluateLayoutPage(parsed, state.Layout, pageBookmark, out value, out status, out message);
                    }
                    return TryEvaluatePage(document, candidate, parsed, out value, out status, out message);
                case WordFieldType.NumChars:
                    return TryFormatIntegerField(ReadFieldStatistic(document, "Characters"), parsed, "Updated from OfficeIMO document statistics Characters.", out value, out status, out message);
                case WordFieldType.NumWords:
                    return TryFormatIntegerField(ReadFieldStatistic(document, "Words"), parsed, "Updated from OfficeIMO document statistics Words.", out value, out status, out message);
                case WordFieldType.NumPages:
                    usedLayout = state.Layout != null;
                    return TryFormatIntegerField(state.Layout?.PageCount ?? (state.EstimatedTotalPages ??= EstimateTotalPages(document)), parsed, state.Layout != null ? "Updated from OfficeIMO page layout." : "Updated from OfficeIMO page-break count.", out value, out status, out message);
                case WordFieldType.RevNum:
                    return TryEvaluateCoreProperty(document, CoreFieldProperty.Revision, parsed, out value, out status, out message);
                case WordFieldType.Section:
                    return TryEvaluateSectionNumber(document, candidate, parsed, out value, out status, out message);
                case WordFieldType.SectionPages:
                    return TryEvaluateSectionPages(document, candidate, parsed, out value, out status, out message);
                case WordFieldType.TOC:
                    status = WordFieldUpdateStatus.Skipped;
                    message = "Table of contents refresh requires Word or another layout-aware application and was left unchanged; enable Settings.UpdateFieldsOnOpen explicitly to delegate document-wide refresh.";
                    return false;
                case WordFieldType.Index:
                    status = WordFieldUpdateStatus.Skipped;
                    message = "Index refresh requires Word or another layout-aware application and was left unchanged; enable Settings.UpdateFieldsOnOpen explicitly to delegate document-wide refresh.";
                    return false;
                default:
                    return false;
            }
        }

        private static bool TryEvaluateInfo(
            WordDocument document,
            WordFieldInventory.ParsedFieldInstruction parsed,
            out string? value,
            out WordFieldUpdateStatus status,
            out string message) {
            value = null;
            status = WordFieldUpdateStatus.Unsupported;
            message = "INFO field is missing a built-in property name.";

            string? propertyName = parsed.Instructions.FirstOrDefault();
            if (string.IsNullOrWhiteSpace(propertyName)) {
                return false;
            }

            propertyName = TrimQuotes(propertyName);
            if (TryEvaluateBuiltInDocumentProperty(document, propertyName, parsed, out value, out status, out message)) {
                return true;
            }

            if (status == WordFieldUpdateStatus.Unsupported) {
                return false;
            }

            status = WordFieldUpdateStatus.Skipped;
            message = $"Built-in document property {propertyName} was not found or is empty.";
            return false;
        }

        private static bool TryEvaluateDocumentProperty(
            WordDocument document,
            WordFieldInventory.ParsedFieldInstruction parsed,
            out string? value,
            out WordFieldUpdateStatus status,
            out string message) {
            value = null;
            status = WordFieldUpdateStatus.Unsupported;
            message = "DOCPROPERTY field is missing a property name.";

            string? propertyName = parsed.Instructions.FirstOrDefault();
            if (string.IsNullOrWhiteSpace(propertyName)) {
                return false;
            }

            propertyName = TrimQuotes(propertyName);

            if (TryEvaluateBuiltInDocumentProperty(document, propertyName, parsed, out value, out status, out message)) {
                return true;
            }

            if (status == WordFieldUpdateStatus.Unsupported) {
                return false;
            }

            KeyValuePair<string, object?> custom = ReadFieldCustomProperty(document, propertyName);
            if (!string.IsNullOrEmpty(custom.Key)) {
                string success = $"Updated from custom document property {custom.Key}.";
                return custom.Value is DateTime date
                    ? TrySetDate(date, parsed, success, out value, out status, out message)
                    : TrySetTextValue(FormatValue(custom.Value), parsed, success, out value, out status, out message);
            }

            status = WordFieldUpdateStatus.Skipped;
            message = $"Document property {propertyName} was not found.";
            return false;
        }

        private static bool TryEvaluateDocumentVariable(
            WordDocument document,
            WordFieldInventory.ParsedFieldInstruction parsed,
            out string? value,
            out WordFieldUpdateStatus status,
            out string message) {
            value = null;
            status = WordFieldUpdateStatus.Unsupported;
            message = "DOCVARIABLE field is missing a variable name.";

            string? variableName = parsed.Instructions.FirstOrDefault();
            if (string.IsNullOrWhiteSpace(variableName)) {
                return false;
            }

            variableName = TrimQuotes(variableName);
            if (string.IsNullOrWhiteSpace(variableName)) {
                return false;
            }

            KeyValuePair<string, string> variable = ReadFieldVariable(document, variableName);

            if (!string.IsNullOrEmpty(variable.Key)) {
                return TrySetTextValue(variable.Value, parsed, $"Updated from document variable {variable.Key}.", out value, out status, out message);
            }

            status = WordFieldUpdateStatus.Skipped;
            message = $"Document variable {variableName} was not found.";
            return false;
        }

        private static bool TryEvaluateQuote(
            string instructionText,
            WordFieldInventory.ParsedFieldInstruction parsed,
            out string? value,
            out WordFieldUpdateStatus status,
            out string message) {
            value = null;
            status = WordFieldUpdateStatus.Unsupported;
            message = "QUOTE fields need exactly one quoted literal instruction to be evaluated by OfficeIMO.";

            if (!TryGetQuoteLiteralAndTail(instructionText, out string? literal, out string tail)) {
                if (parsed.Instructions.Count == 1) {
                    literal = parsed.Instructions[0].Trim();
                    tail = string.Empty;
                } else {
                    return false;
                }
            }

            WordFieldInventory.ParsedFieldInstruction switchParsed = parsed;
            if (!string.IsNullOrWhiteSpace(tail)) {
                switchParsed = WordFieldInventory.ParseInstruction("QUOTE \"OfficeIMO\" " + tail);
                if (switchParsed.Diagnostics.Count > 0 || switchParsed.FieldType != WordFieldType.Quote) {
                    message = string.Join(" ", switchParsed.Diagnostics.DefaultIfEmpty("QUOTE field switches could not be parsed."));
                    return false;
                }
            } else {
                switchParsed = WordFieldInventory.ParseInstruction("QUOTE \"OfficeIMO\"");
            }

            if (string.IsNullOrWhiteSpace(literal)) {
                return false;
            }

            string quoteLiteral = literal!;
            if (switchParsed.Switches.Count > 0) {
                message = "QUOTE fields with general switches are not evaluated by OfficeIMO.";
                return false;
            }

            if (quoteLiteral.Length < 2 || quoteLiteral[0] != '"' || quoteLiteral[quoteLiteral.Length - 1] != '"') {
                message = "QUOTE fields need a quoted literal instruction to be evaluated by OfficeIMO.";
                return false;
            }

            string literalValue = TrimQuotes(quoteLiteral);
            if (!string.IsNullOrWhiteSpace(switchParsed.NumericPictureSwitch)) {
                if (GetLastMeaningfulFormat(switchParsed.FormatSwitches) != null) {
                    message = @"QUOTE fields cannot combine \# numeric picture and \* format switches for deterministic literal refresh.";
                    return false;
                }

                if (!decimal.TryParse(literalValue, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal numericLiteral)) {
                    message = "QUOTE field numeric picture switch requires a numeric literal.";
                    return false;
                }

                if (!TryFormatFormulaValue(numericLiteral, switchParsed.NumericPictureSwitch, out string numericPictureValue, out string? numericPictureDiagnostic)) {
                    message = numericPictureDiagnostic == null
                        ? "QUOTE field numeric picture switch could not be applied."
                        : numericPictureDiagnostic.Replace("Formula numeric picture", "Field numeric picture");
                    return false;
                }

                value = numericPictureValue;
                status = WordFieldUpdateStatus.Updated;
                message = "Updated from QUOTE numeric literal instruction with numeric picture formatting.";
                return true;
            }

            if (!TryApplyQuoteFormat(switchParsed.FormatSwitches, literalValue, out string formattedValue, out string? unsupportedFormat, out bool usedNumericFormat)) {
                message = $"QUOTE field format switch {unsupportedFormat} is not supported for literal refresh.";
                return false;
            }

            value = formattedValue;
            status = WordFieldUpdateStatus.Updated;
            message = parsed.FormatSwitches.Any(fieldFormat => fieldFormat != WordFieldFormat.MergeFormat && fieldFormat != WordFieldFormat.CharFormat)
                ? usedNumericFormat
                    ? "Updated from QUOTE numeric literal instruction with deterministic number formatting."
                    : "Updated from QUOTE literal instruction with deterministic text formatting."
                : "Updated from QUOTE literal instruction.";
            return true;
        }

        private static bool TryApplyQuoteFormat(
            IReadOnlyList<WordFieldFormat> formatSwitches,
            string source,
            out string value,
            out string? unsupportedFormat,
            out bool usedNumericFormat) {
            usedNumericFormat = false;
            WordFieldFormat? format = GetLastMeaningfulFormat(formatSwitches);
            switch (format) {
                case WordFieldFormat.UpperRoman:
                case WordFieldFormat.LowerRoman:
                case WordFieldFormat.Arabic:
                case WordFieldFormat.Ordinal:
                case WordFieldFormat.LowerLetter:
                case WordFieldFormat.UpperLetter:
                case WordFieldFormat.Hex:
                case WordFieldFormat.CardText:
                case WordFieldFormat.OrdText:
                case WordFieldFormat.DollarText:
                    if (!int.TryParse(source, NumberStyles.Integer, CultureInfo.InvariantCulture, out int numericValue)) {
                        value = source;
                        unsupportedFormat = format.Value.ToString();
                        return false;
                    }

                    if (RequiresNonNegativeNumber(format.Value) && numericValue < 0) {
                        value = source;
                        unsupportedFormat = format.Value.ToString();
                        return false;
                    }

                    value = FormatSequenceValue(numericValue, new[] { format.Value });
                    unsupportedFormat = null;
                    usedNumericFormat = true;
                    return true;
                default:
                    usedNumericFormat = false;
                    return TryApplyReferenceTextFormat(formatSwitches, source, out value, out unsupportedFormat);
            }
        }

        private static bool RequiresNonNegativeNumber(WordFieldFormat format) {
            return format == WordFieldFormat.Hex ||
                format == WordFieldFormat.CardText ||
                format == WordFieldFormat.OrdText ||
                format == WordFieldFormat.DollarText;
        }

        private static IReadOnlyList<MutableFieldCandidate> SortFieldsForEvaluation(IReadOnlyList<MutableFieldCandidate> candidates) {
            var instructionChildren = candidates
                .Where(candidate => candidate.InstructionParentSequence.HasValue)
                .GroupBy(candidate => candidate.InstructionParentSequence!.Value)
                .ToDictionary(group => group.Key, group => group.OrderBy(candidate => candidate.Sequence).ToArray());

            if (instructionChildren.Count == 0) return candidates;
            var ordered = new List<MutableFieldCandidate>(candidates.Count);
            var visited = new HashSet<int>();
            var pending = new Stack<(MutableFieldCandidate Candidate, bool Ready)>();
            foreach (MutableFieldCandidate candidate in candidates) {
                if (visited.Contains(candidate.Sequence)) continue;
                pending.Push((candidate, false));
                while (pending.Count != 0) {
                    var (current, ready) = pending.Pop();
                    if (ready) {
                        ordered.Add(current);
                    } else if (visited.Add(current.Sequence)) {
                        pending.Push((current, true));
                        if (instructionChildren.TryGetValue(current.Sequence, out var children)) {
                            for (int index = children.Length - 1; index >= 0; index--) pending.Push((children[index], false));
                        }
                    }
                }
            }
            return ordered;
        }

        private static bool TryEvaluateBuiltInDocumentProperty(
            WordDocument document,
            string propertyName,
            WordFieldInventory.ParsedFieldInstruction parsed,
            out string? value,
            out WordFieldUpdateStatus status,
            out string message) {
            value = null;
            status = WordFieldUpdateStatus.Skipped;
            message = string.Empty;

            switch (propertyName.Replace(" ", string.Empty).ToUpperInvariant()) {
                case "AUTHOR":
                case "CREATOR":
                    return TryEvaluateCoreProperty(document, CoreFieldProperty.Creator, parsed, out value, out status, out message);
                case "TITLE":
                    return TryEvaluateCoreProperty(document, CoreFieldProperty.Title, parsed, out value, out status, out message);
                case "SUBJECT":
                    return TryEvaluateCoreProperty(document, CoreFieldProperty.Subject, parsed, out value, out status, out message);
                case "CATEGORY":
                    return TryEvaluateCoreProperty(document, CoreFieldProperty.Category, parsed, out value, out status, out message);
                case "KEYWORDS":
                    return TryEvaluateCoreProperty(document, CoreFieldProperty.Keywords, parsed, out value, out status, out message);
                case "COMMENTS":
                case "DESCRIPTION":
                    return TryEvaluateCoreProperty(document, CoreFieldProperty.Description, parsed, out value, out status, out message);
                case "LASTSAVEDBY":
                case "LASTMODIFIEDBY":
                    return TryEvaluateCoreProperty(document, CoreFieldProperty.LastModifiedBy, parsed, out value, out status, out message);
                case "LASTPRINTED":
                case "PRINTDATE":
                    return TryEvaluateCoreProperty(document, CoreFieldProperty.LastPrinted, parsed, out value, out status, out message);
                case "CREATED":
                case "CREATEDATE":
                    return TryEvaluateCoreProperty(document, CoreFieldProperty.Created, parsed, out value, out status, out message);
                case "MODIFIED":
                case "SAVEDATE":
                    return TryEvaluateCoreProperty(document, CoreFieldProperty.Modified, parsed, out value, out status, out message);
                case "REVISION":
                case "REVNUM":
                    return TryEvaluateCoreProperty(document, CoreFieldProperty.Revision, parsed, out value, out status, out message);
                case "VERSION":
                    return TryEvaluateCoreProperty(document, CoreFieldProperty.Version, parsed, out value, out status, out message);
                default:
                    return false;
            }
        }

        private static bool TryEvaluateFileName(
            WordDocument document,
            WordFieldInventory.ParsedFieldInstruction parsed,
            out string? value,
            out WordFieldUpdateStatus status,
            out string message) {
            value = null;
            status = WordFieldUpdateStatus.Skipped;

            string? filePath = document.FilePath;
            if (string.IsNullOrWhiteSpace(filePath)) {
                message = "FILENAME cannot be evaluated because the document has no backing file path.";
                return false;
            }

            bool includePath = parsed.Switches.Any(fieldSwitch => string.Equals(fieldSwitch, "\\p", StringComparison.OrdinalIgnoreCase));
            string source = includePath ? filePath! : Path.GetFileName(filePath!);
            if (!TryApplyReferenceTextFormat(parsed.FormatSwitches, source, out string formattedValue, out string? unsupportedFormat)) {
                value = null;
                status = WordFieldUpdateStatus.Unsupported;
                message = $"FILENAME format switch {unsupportedFormat} is not supported for deterministic refresh.";
                return false;
            }

            value = formattedValue;
            status = WordFieldUpdateStatus.Updated;
            message = includePath ? "Updated from document file path." : "Updated from document file name.";
            if (GetLastMeaningfulFormat(parsed.FormatSwitches) != null) {
                message += " Text format switch was applied.";
            }

            return true;
        }

        private static bool TryGetQuoteLiteralAndTail(string instructionText, out string? literal, out string tail) {
            literal = null;
            tail = string.Empty;
            if (string.IsNullOrWhiteSpace(instructionText)) {
                return false;
            }

            string text = instructionText.Trim();
            if (!text.StartsWith("QUOTE", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            int index = 5;
            while (index < text.Length && char.IsWhiteSpace(text[index])) {
                index++;
            }

            if (index >= text.Length || text[index] != '"') {
                return false;
            }

            int start = index;
            index++;
            while (index < text.Length) {
                if (text[index] == '"') {
                    literal = text.Substring(start, index - start + 1);
                    tail = index + 1 < text.Length ? text.Substring(index + 1).Trim() : string.Empty;
                    return true;
                }

                index++;
            }

            return false;
        }

        private static bool TryEvaluateFileSize(
            WordDocument document,
            WordFieldInventory.ParsedFieldInstruction parsed,
            out string? value,
            out WordFieldUpdateStatus status,
            out string message) {
            value = null;
            status = WordFieldUpdateStatus.Skipped;

            if (string.IsNullOrWhiteSpace(document.FilePath)) {
                message = "FILESIZE cannot be evaluated because the document has no backing file path.";
                return false;
            }

            if (!File.Exists(document.FilePath)) {
                message = "FILESIZE cannot be evaluated because the backing file does not exist.";
                return false;
            }

            if (!TryGetFileSizeUnit(parsed.Switches, out FileSizeUnit unit, out string? unsupportedSwitch)) {
                status = WordFieldUpdateStatus.Unsupported;
                message = $"FILESIZE switch {unsupportedSwitch} is not supported for deterministic refresh.";
                return false;
            }

            long bytes = new FileInfo(document.FilePath).Length;
            decimal numericValue = unit switch {
                FileSizeUnit.Kilobytes => Math.Round(bytes / 1000m, 0, MidpointRounding.AwayFromZero),
                FileSizeUnit.Megabytes => Math.Round(bytes / 1000000m, 0, MidpointRounding.AwayFromZero),
                _ => bytes
            };

            string unitMessage = unit switch {
                FileSizeUnit.Kilobytes => "rounded decimal kilobytes",
                FileSizeUnit.Megabytes => "rounded decimal megabytes",
                _ => "bytes"
            };

            return TryFormatIntegerField(numericValue, parsed, $"Updated from backing DOCX package file size in {unitMessage}.", out value, out status, out message);
        }

        private static bool TryGetFileSizeUnit(IReadOnlyList<string> switches, out FileSizeUnit unit, out string? unsupportedSwitch) {
            unit = FileSizeUnit.Bytes;
            unsupportedSwitch = null;

            foreach (string fieldSwitch in switches) {
                string normalizedSwitch = fieldSwitch.Trim();
                if (string.Equals(normalizedSwitch, "\\k", StringComparison.OrdinalIgnoreCase)) {
                    unit = FileSizeUnit.Kilobytes;
                    continue;
                }

                if (string.Equals(normalizedSwitch, "\\m", StringComparison.OrdinalIgnoreCase)) {
                    unit = FileSizeUnit.Megabytes;
                    continue;
                }

                unsupportedSwitch = normalizedSwitch;
                return false;
            }

            return true;
        }

        private static bool TryEvaluatePage(
            WordDocument document,
            MutableFieldCandidate candidate,
            WordFieldInventory.ParsedFieldInstruction parsed,
            out string? value,
            out WordFieldUpdateStatus status,
            out string message) {
            value = null;
            status = WordFieldUpdateStatus.Skipped;

            if (candidate.LocationKind != WordFieldLocationKind.Body) {
                message = "PAGE fields outside the document body need Word layout context and were left unchanged.";
                return false;
            }

            int? page = EstimatePageForBodyField(document, candidate.AnchorElement);
            if (page == null) {
                message = "PAGE field position could not be matched to a body paragraph.";
                return false;
            }

            return TryFormatIntegerField(page.Value, parsed, "Updated from OfficeIMO page-break order.", out value, out status, out message);
        }

        private static bool TryFormatNumericField(
            decimal numericValue,
            string? numericPicture,
            string successMessage,
            out string? value,
            out WordFieldUpdateStatus status,
            out string message) {
            value = null;
            if (!TryFormatFormulaValue(numericValue, numericPicture, out string formattedValue, out string? diagnostic)) {
                status = WordFieldUpdateStatus.Unsupported;
                message = diagnostic == null
                    ? "Field numeric picture switch could not be applied."
                    : diagnostic.Replace("Formula numeric picture", "Field numeric picture");
                return false;
            }

            value = formattedValue;
            status = WordFieldUpdateStatus.Updated;
            message = string.IsNullOrWhiteSpace(numericPicture)
                ? successMessage
                : $"{successMessage} Numeric picture formatting was applied.";
            return true;
        }

        private static bool TryFormatIntegerField(
            decimal numericValue,
            WordFieldInventory.ParsedFieldInstruction parsed,
            string successMessage,
            out string? value,
            out WordFieldUpdateStatus status,
            out string message) {
            value = null;
            WordFieldFormat? format = GetLastMeaningfulFormat(parsed.FormatSwitches);

            if (!string.IsNullOrWhiteSpace(parsed.NumericPictureSwitch) && format != null) {
                status = WordFieldUpdateStatus.Unsupported;
                message = "Field cannot combine numeric picture and general format switches for deterministic refresh.";
                return false;
            }

            if (!string.IsNullOrWhiteSpace(parsed.NumericPictureSwitch)) {
                return TryFormatNumericField(numericValue, parsed.NumericPictureSwitch, successMessage, out value, out status, out message);
            }

            switch (format) {
                case null:
                case WordFieldFormat.Arabic:
                    value = numericValue.ToString("0", CultureInfo.InvariantCulture);
                    status = WordFieldUpdateStatus.Updated;
                    message = successMessage;
                    return true;
                case WordFieldFormat.UpperRoman:
                case WordFieldFormat.LowerRoman:
                case WordFieldFormat.Ordinal:
                case WordFieldFormat.LowerLetter:
                case WordFieldFormat.UpperLetter:
                case WordFieldFormat.Hex:
                case WordFieldFormat.CardText:
                case WordFieldFormat.OrdText:
                case WordFieldFormat.DollarText:
                    if (numericValue != decimal.Truncate(numericValue) || numericValue < int.MinValue || numericValue > int.MaxValue) {
                        status = WordFieldUpdateStatus.Unsupported;
                        message = $"Field format switch {format.Value} requires an integer value in the deterministic refresh range.";
                        return false;
                    }

                    int sequenceValue = decimal.ToInt32(numericValue);
                    if (RequiresNonNegativeNumber(format.Value) && sequenceValue < 0) {
                        status = WordFieldUpdateStatus.Unsupported;
                        message = $"Field format switch {format.Value} requires a non-negative value for deterministic numeric refresh.";
                        return false;
                    }

                    value = FormatSequenceValue(sequenceValue, new[] { format.Value });
                    status = WordFieldUpdateStatus.Updated;
                    message = successMessage + " General numeric format switch was applied.";
                    return true;
                default:
                    status = WordFieldUpdateStatus.Unsupported;
                    message = $"Field format switch {format.Value} is not supported for deterministic numeric refresh.";
                    return false;
            }
        }

        private enum FileSizeUnit {
            Bytes,
            Kilobytes,
            Megabytes
        }

        private static bool TrySetValue(string? source, string successMessage, out string? value, out WordFieldUpdateStatus status, out string message) {
            value = source;
            if (!string.IsNullOrEmpty(value)) {
                status = WordFieldUpdateStatus.Updated;
                message = successMessage;
                return true;
            }

            status = WordFieldUpdateStatus.Skipped;
            message = "Source document property is empty.";
            return false;
        }

        private static bool TrySetTextValue(
            string? source,
            WordFieldInventory.ParsedFieldInstruction parsed,
            string successMessage,
            out string? value,
            out WordFieldUpdateStatus status,
            out string message) {
            if (string.IsNullOrEmpty(source)) {
                value = null;
                status = WordFieldUpdateStatus.Skipped;
                message = "Source document property is empty.";
                return false;
            }

            string sourceText = source!;
            if (!TryApplyReferenceTextFormat(parsed.FormatSwitches, sourceText, out string formattedValue, out string? unsupportedFormat)) {
                value = null;
                status = WordFieldUpdateStatus.Unsupported;
                message = $"Field format switch {unsupportedFormat} is not supported for deterministic metadata refresh.";
                return false;
            }

            value = formattedValue;
            status = WordFieldUpdateStatus.Updated;
            message = GetLastMeaningfulFormat(parsed.FormatSwitches) == null
                ? successMessage
                : successMessage + " Text format switch was applied.";
            return true;
        }

        private static bool TrySetDate(DateTime? source, WordFieldInventory.ParsedFieldInstruction parsed, string successMessage, out string? value, out WordFieldUpdateStatus status, out string message) {
            if (!source.HasValue) {
                value = null;
                status = WordFieldUpdateStatus.Skipped;
                message = "Source document property is empty.";
                return false;
            }

            if (!TryFormatDateTime(source.Value, parsed, out value, out message)) {
                status = WordFieldUpdateStatus.Unsupported;
                return false;
            }

            if (!TryApplyReferenceTextFormat(parsed.FormatSwitches, value, out string formattedValue, out string? unsupportedFormat)) {
                status = WordFieldUpdateStatus.Unsupported;
                message = $"Text format switch {unsupportedFormat} is not supported for deterministic date refresh.";
                return false;
            }

            value = formattedValue;
            if (!string.IsNullOrEmpty(value)) {
                status = WordFieldUpdateStatus.Updated;
                message = GetLastMeaningfulFormat(parsed.FormatSwitches) == null
                    ? successMessage
                    : successMessage + " Text format switch was applied.";
                return true;
            }

            status = WordFieldUpdateStatus.Skipped;
            message = "Source document property is empty.";
            return false;
        }

        private static string FormatValue(object? value) {
            if (value == null) {
                return string.Empty;
            }

            if (value is DateTime dateTime) {
                return dateTime.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            }

            return OfficeIMO.OfficeRoundTripNumberText.FormatObject(value, CultureInfo.InvariantCulture) ?? string.Empty;
        }

        private static string TrimQuotes(string value) {
            value = value.Trim();
            return value.Length >= 2 && value[0] == '"' && value[value.Length - 1] == '"'
                ? value.Substring(1, value.Length - 2)
                : value;
        }

        private static void SetResultText(WordDocument document, MutableFieldCandidate candidate, string value) {
            WordPackageModel? model = document._ownedModel;
            if (model != null) {
                WordOwnedEdits.PrepareBodyEdit(model);
                model.MarkChanged();
            }
            const int W = FieldNamespace;
            if (candidate.SimpleField.Exists) {
                List<OfficeXmlNode> simpleTexts = WordFieldInventory.EnumerateFieldContent(candidate.SimpleField)
                    .Where(element => element.Is(W, "t")).ToList();
                if (simpleTexts.Count > 0) SetText(simpleTexts, value);
                else candidate.SimpleField.Append(CreateResultRun(candidate.SimpleField, value));
                return;
            }

            var texts = GetResultTexts(candidate).ToList();
            if (texts.Count > 0) {
                SetText(texts, value);
                return;
            }

            OfficeXmlNode reference = candidate.EndRun.Exists ? candidate.EndRun : candidate.AnchorElement;
            OfficeXmlNode parent = reference.Parent;
            OfficeXmlNode run = CreateResultRun(parent, value);
            if (candidate.EndRun.Exists) {
                if (!candidate.HasSeparator) {
                    OfficeXmlNode separatorRun = parent.CreateElement(W, "r");
                    OfficeXmlNode marker = separatorRun.CreateElement(W, "fldChar");
                    marker.SetAttribute(W, "fldCharType", "separate");
                    separatorRun.Append(marker);
                    parent.InsertBefore(separatorRun, reference);
                }
                run = parent.InsertBefore(run, reference);
            } else {
                // A malformed field with no end marker still accepts a cached result.
                run = reference.InsertAfterSelf(run);
            }
            candidate.ResultRuns.Add(run);
        }

        private static IEnumerable<OfficeXmlNode> GetResultTexts(MutableFieldCandidate candidate) {
            const int W = FieldNamespace;
            foreach (OfficeXmlNode run in candidate.ResultRuns) {
                if (run != candidate.EndRun) {
                    foreach (OfficeXmlNode text in run.Children(W, "t")) yield return text;
                    continue;
                }

                foreach (OfficeXmlNode child in run.Children()) {
                    if (child.Is(W, "fldChar") && child.Attribute(W, "fldCharType") == "end") break;
                    if (child.Is(W, "t")) yield return child;
                }
            }
        }

        private static OfficeXmlNode CreateResultRun(OfficeXmlNode parent, string value) {
            OfficeXmlNode run = parent.CreateElement(FieldNamespace, "r");
            OfficeXmlNode text = run.CreateElement(FieldNamespace, "t");
            SetText(text, value);
            run.Append(text);
            return run;
        }

        private static void SetText(IReadOnlyList<OfficeXmlNode> texts, string value) {
            for (int index = 0; index < texts.Count; index++) SetText(texts[index], index == 0 ? value : string.Empty);
        }

        private static void SetText(OfficeXmlNode text, string value) {
            text.SetInnerText(value);
            text.SetAttribute(OfficeXmlNamespaces.Xml, "space",
                value.Length > 0 && (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[value.Length - 1])) ? "preserve" : null);
        }

        private static List<MutableFieldCandidate> EnumerateFields(WordDocument document) {
            const int W = FieldNamespace;
            var candidates = new List<MutableFieldCandidate>();
            foreach (WordFieldInventory.NodeFieldRoot story in WordFieldInventory.EnumerateStories(document)) {
                var stack = new Stack<ComplexFieldBuilder>();
                foreach (OfficeXmlNode element in WordFieldInventory.EnumerateFieldContent(story.Root)) {
                    if (element.Is(W, "fldSimple")) {
                        int nestingLevel = stack.Count;
                        for (OfficeXmlNode parent = element.Parent; parent.Exists; parent = parent.Parent) {
                            if (parent.Is(W, "fldSimple")) nestingLevel++;
                        }
                        candidates.Add(MutableFieldCandidate.ForSimple(story, candidates.Count, nestingLevel, element));
                    } else if (element.Is(W, "r")) {
                        ProcessRun(story, element, stack, candidates);
                    }
                }

                while (stack.Count > 0) {
                    ComplexFieldBuilder builder = stack.Pop();
                    candidates[builder.Sequence] = builder.ToCandidate(story);
                }
            }
            return candidates;
        }

        private static bool IsNestedInsideReplacedField(MutableFieldCandidate candidate, IEnumerable<MutableFieldCandidate> replacedContainingFields) {
            if (candidate.NestingLevel == 0) return false;
            foreach (MutableFieldCandidate containingField in replacedContainingFields) {
                if (containingField.EndRun.IsNone || candidate.Sequence <= containingField.Sequence ||
                    candidate.NestingLevel <= containingField.NestingLevel || candidate.PartUri != containingField.PartUri) continue;
                if (IsBefore(containingField.AnchorElement, candidate.AnchorElement) &&
                    IsBefore(candidate.AnchorElement, containingField.EndRun)) return true;
            }
            return false;
        }

        private static bool IsBefore(OfficeXmlNode first, OfficeXmlNode second) {
            if (first.IsNone || second.IsNone || first == second) return false;
            var ancestors = new HashSet<OfficeXmlNode>();
            for (OfficeXmlNode node = second; node.Exists; node = node.Parent) ancestors.Add(node);
            if (ancestors.Contains(first)) return true;
            for (OfficeXmlNode node = first; node.Exists && !ancestors.Contains(node); node = node.Parent) {
                for (OfficeXmlNode sibling = node.NextSibling; sibling.Exists; sibling = sibling.NextSibling) {
                    if (ancestors.Contains(sibling)) return true;
                }
            }
            return false;
        }

        private static void ProcessRun(WordFieldInventory.NodeFieldRoot story, OfficeXmlNode run,
            Stack<ComplexFieldBuilder> stack, List<MutableFieldCandidate> candidates) {
            const int W = FieldNamespace;
            bool resultRunRecorded = false;
            foreach (OfficeXmlNode child in run.Children()) {
                if (child.Is(W, "fldChar")) {
                    string? fieldCharType = child.Attribute(W, "fldCharType");
                    if (fieldCharType == "begin") {
                        stack.Push(new ComplexFieldBuilder(candidates.Count, stack.Count, run) {
                            IsLocked = WordFieldInventory.ReadFlag(child, "fldLock")
                        });
                        candidates.Add(null!);
                        continue;
                    }

                    if (stack.Count == 0) continue;
                    if (fieldCharType == "separate") {
                        stack.Peek().HasSeparator = true;
                    } else if (fieldCharType == "end") {
                        ComplexFieldBuilder completed = stack.Pop();
                        completed.EndRun = run;
                        if (stack.Count > 0 && !stack.Peek().HasSeparator) {
                            ComplexFieldBuilder parent = stack.Peek();
                            completed.InstructionParentSequence = parent.Sequence;
                            parent.InstructionParts.Add(InstructionPart.ForResultRuns(completed.ResultRuns));
                        }
                        candidates[completed.Sequence] = completed.ToCandidate(story);
                    }
                    continue;
                }

                if (stack.Count == 0) continue;
                if (child.Is(W, "instrText")) {
                    stack.Peek().InstructionParts.Add(InstructionPart.ForLiteral(child.InnerText));
                } else if (!resultRunRecorded && child.Is(W, "t")) {
                    foreach (ComplexFieldBuilder builder in stack) {
                        if (builder.HasSeparator) builder.ResultRuns.Add(run);
                    }
                    resultRunRecorded = true;
                }
            }
        }

        private sealed class MutableFieldCandidate {
            private MutableFieldCandidate(
                int sequence,
                WordFieldRepresentation representation,
                WordFieldLocationKind locationKind,
                string partUri,
                IReadOnlyList<InstructionPart> instructionParts,
                int nestingLevel,
                int? instructionParentSequence,
                OfficeXmlNode anchorElement,
                OfficeXmlNode simpleField,
                List<OfficeXmlNode> resultRuns,
                OfficeXmlNode endRun,
                bool hasSeparator,
                bool isLocked) {
                Sequence = sequence;
                Representation = representation;
                LocationKind = locationKind;
                PartUri = partUri;
                _instructionParts = instructionParts;
                NestingLevel = nestingLevel;
                InstructionParentSequence = instructionParentSequence;
                AnchorElement = anchorElement;
                SimpleField = simpleField;
                ResultRuns = resultRuns;
                EndRun = endRun;
                HasSeparator = hasSeparator;
                IsLocked = isLocked;
            }

            internal int Sequence { get; }

            internal int Index { get; set; }

            internal WordFieldRepresentation Representation { get; }

            internal WordFieldLocationKind LocationKind { get; }

            internal string PartUri { get; }

            private readonly IReadOnlyList<InstructionPart> _instructionParts;

            internal string InstructionText => string.Concat(_instructionParts.Select(part => part.GetText()));

            internal bool HasDynamicInstruction => _instructionParts.Any(part => part.IsDynamic);

            internal int NestingLevel { get; }

            internal int? InstructionParentSequence { get; }

            internal OfficeXmlNode AnchorElement { get; }


            internal OfficeXmlNode SimpleField { get; }

            internal List<OfficeXmlNode> ResultRuns { get; }

            internal OfficeXmlNode EndRun { get; }

            internal bool HasSeparator { get; }

            internal bool IsLocked { get; }

            internal static MutableFieldCandidate ForSimple(WordFieldInventory.NodeFieldRoot root, int sequence, int nestingLevel, OfficeXmlNode simpleField) {
                return new MutableFieldCandidate(
                    sequence,
                    WordFieldRepresentation.Simple,
                    root.LocationKind,
                    root.PartUri,
                    new[] { InstructionPart.ForLiteral(simpleField.Attribute(FieldNamespace, "instr") ?? string.Empty) },
                    nestingLevel,
                    null,
                    simpleField,
                    simpleField,
                    simpleField.Children(FieldNamespace, "r").ToList(),
                    default,
                    true,
                    WordFieldInventory.ReadFlag(simpleField, "fldLock"));
            }

            internal WordFieldUpdateResult ToResult(
                WordFieldType? fieldType,
                WordFieldUpdateStatus status,
                string? resultText,
                string message,
                WordFieldEvaluationReason reason = WordFieldEvaluationReason.Default) {
                return new WordFieldUpdateResult(
                    Index,
                    Representation,
                    LocationKind,
                    PartUri,
                    InstructionText,
                    fieldType,
                    status,
                    resultText,
                    message,
                    IsLocked,
                    reason);
            }

            internal static MutableFieldCandidate ForComplex(WordFieldInventory.NodeFieldRoot root, ComplexFieldBuilder builder) {
                return new MutableFieldCandidate(
                    builder.Sequence,
                    WordFieldRepresentation.Complex,
                    root.LocationKind,
                    root.PartUri,
                    builder.InstructionParts.ToArray(),
                    builder.NestingLevel,
                    builder.InstructionParentSequence,
                    builder.AnchorElement,
                    default,
                    builder.ResultRuns,
                    builder.EndRun,
                    builder.HasSeparator,
                    builder.IsLocked);
            }
        }

        private sealed class ComplexFieldBuilder {
            internal ComplexFieldBuilder(int sequence, int nestingLevel, OfficeXmlNode anchorElement) {
                Sequence = sequence;
                NestingLevel = nestingLevel;
                AnchorElement = anchorElement;
            }

            internal int Sequence { get; }

            internal int NestingLevel { get; }

            internal OfficeXmlNode AnchorElement { get; }

            internal List<InstructionPart> InstructionParts { get; } = new();

            internal List<OfficeXmlNode> ResultRuns { get; } = new();

            internal int? InstructionParentSequence { get; set; }

            internal bool HasSeparator { get; set; }

            internal bool IsLocked { get; set; }

            internal OfficeXmlNode EndRun { get; set; }

            internal MutableFieldCandidate ToCandidate(WordFieldInventory.NodeFieldRoot root) {
                return MutableFieldCandidate.ForComplex(root, this);
            }
        }

        private sealed class InstructionPart {
            private readonly string? _literal;
            private readonly IReadOnlyList<OfficeXmlNode>? _resultRuns;

            private InstructionPart(string literal) {
                _literal = literal;
            }

            private InstructionPart(IReadOnlyList<OfficeXmlNode> resultRuns) {
                _resultRuns = resultRuns;
            }

            internal bool IsDynamic => _resultRuns != null;

            internal static InstructionPart ForLiteral(string literal) => new InstructionPart(literal);

            internal static InstructionPart ForResultRuns(IReadOnlyList<OfficeXmlNode> resultRuns) => new InstructionPart(resultRuns);

            internal string GetText() {
                if (_literal != null) {
                    return _literal;
                }

                if (_resultRuns == null) return string.Empty;
                var text = new StringBuilder();
                foreach (OfficeXmlNode run in _resultRuns) {
                    foreach (OfficeXmlNode result in run.Children(FieldNamespace, "t")) text.Append(result.InnerText);
                }
                return text.ToString();
            }
        }

        private sealed class FieldEvaluationState {
            internal Dictionary<string, int> Sequences { get; } = new(StringComparer.OrdinalIgnoreCase);

            internal Dictionary<string, OfficeXmlNode> SequenceHeadingResetKeys { get; } = new(StringComparer.OrdinalIgnoreCase);

            internal Dictionary<string, ListNumSequenceState> ListNumSequences { get; } = new(StringComparer.OrdinalIgnoreCase);

            internal Dictionary<int, ReferenceNumberingDefinition>? ReferenceNumberingDefinitions { get; set; }

            internal IReadOnlyDictionary<string, ReferenceParagraphNumbering>? ReferenceStyleNumbering { get; set; }

            internal WordLayoutPagination? Layout { get; set; }
            internal int? EstimatedTotalPages { get; set; }

            internal Dictionary<OfficeXmlNode, string> PageFieldBookmarks { get; } = new();

            internal HashSet<OfficeXmlNode> RegeneratedTablesOfContents { get; } = new();
            internal List<MutableFieldCandidate> LockedTocFields { get; } = new();
        }
    }
}
