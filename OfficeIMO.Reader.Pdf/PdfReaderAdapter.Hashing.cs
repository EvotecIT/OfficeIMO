namespace OfficeIMO.Reader.Pdf;

internal static partial class PdfReaderAdapter {
    private static string BuildChunkMetadataHashInput(ReaderChunk chunk) {
        var builder = new StringBuilder();
        AppendTablesHashInput(builder, chunk.Tables);
        AppendVisualsHashInput(builder, chunk.Visuals);
        AppendFormFieldsHashInput(builder, chunk.FormFields);
        AppendActionsHashInput(builder, chunk.Actions);
        AppendDiagnosticsHashInput(builder, chunk.Diagnostics);
        return builder.ToString();
    }

    private static void AppendTablesHashInput(StringBuilder builder, IReadOnlyList<ReaderTable>? tables) {
        AppendHashValue(builder, "tables.count", tables?.Count ?? 0);
        if (tables is null) return;

        for (int i = 0; i < tables.Count; i++) {
            ReaderTable table = tables[i];
            AppendHashValue(builder, "table.index", i);
            AppendHashValue(builder, "table.title", table.Title);
            AppendHashValue(builder, "table.kind", table.Kind);
            AppendHashValue(builder, "table.callId", table.CallId);
            AppendHashValue(builder, "table.summary", table.Summary);
            AppendHashValue(builder, "table.payloadHash", table.PayloadHash);
            AppendLocationHashInput(builder, "table.location", table.Location);
            AppendStringListHashInput(builder, "table.columns", table.Columns);
            AppendHashValue(builder, "table.totalRowCount", table.TotalRowCount);
            AppendHashValue(builder, "table.truncated", table.Truncated);
            AppendHashValue(builder, "table.rows.count", table.Rows.Count);
            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
                AppendStringListHashInput(builder, "table.row." + rowIndex.ToString(CultureInfo.InvariantCulture), table.Rows[rowIndex]);
            }

            AppendHashValue(builder, "table.columnProfiles.count", table.ColumnProfiles.Count);
            for (int profileIndex = 0; profileIndex < table.ColumnProfiles.Count; profileIndex++) {
                ReaderTableColumnProfile profile = table.ColumnProfiles[profileIndex];
                AppendHashValue(builder, "table.profile.index", profile.Index);
                AppendHashValue(builder, "table.profile.name", profile.Name);
                AppendHashValue(builder, "table.profile.kind", profile.Kind);
                AppendHashValue(builder, "table.profile.nonEmpty", profile.NonEmptyCellCount);
                AppendHashValue(builder, "table.profile.numeric", profile.NumericCellCount);
                AppendHashValue(builder, "table.profile.confidence", profile.Confidence);
            }

            ReaderTableDiagnostics? diagnostics = table.Diagnostics;
            AppendHashValue(builder, "table.diagnostics.hasValue", diagnostics is not null);
            if (diagnostics is not null) {
                AppendHashValue(builder, "table.diagnostics.confidence", diagnostics.Confidence);
                AppendHashValue(builder, "table.diagnostics.schema", diagnostics.SchemaConfidence);
                AppendHashValue(builder, "table.diagnostics.completeness", diagnostics.CellCompleteness);
                AppendHashValue(builder, "table.diagnostics.columnGeometry", diagnostics.ColumnGeometryConfidence);
                AppendHashValue(builder, "table.diagnostics.sourceRows", diagnostics.SourceRowCount);
                AppendHashValue(builder, "table.diagnostics.expectedCells", diagnostics.ExpectedCellCount);
                AppendHashValue(builder, "table.diagnostics.filledCells", diagnostics.FilledCellCount);
                AppendHashValue(builder, "table.diagnostics.missingCells", diagnostics.MissingCellCount);
                AppendHashValue(builder, "table.diagnostics.xStart", diagnostics.XStart);
                AppendHashValue(builder, "table.diagnostics.xEnd", diagnostics.XEnd);
                AppendHashValue(builder, "table.diagnostics.yTop", diagnostics.YTop);
                AppendHashValue(builder, "table.diagnostics.yBottom", diagnostics.YBottom);
                AppendHashValue(builder, "table.diagnostics.width", diagnostics.Width);
                AppendHashValue(builder, "table.diagnostics.height", diagnostics.Height);
                AppendHashValue(builder, "table.diagnostics.hasGeometry", diagnostics.HasGeometry);
            }
        }
    }

    private static void AppendVisualsHashInput(StringBuilder builder, IReadOnlyList<ReaderVisual>? visuals) {
        AppendHashValue(builder, "visuals.count", visuals?.Count ?? 0);
        if (visuals is null) return;

        for (int i = 0; i < visuals.Count; i++) {
            ReaderVisual visual = visuals[i];
            AppendHashValue(builder, "visual.index", i);
            AppendHashValue(builder, "visual.kind", visual.Kind);
            AppendHashValue(builder, "visual.language", visual.Language);
            AppendHashValue(builder, "visual.content", visual.Content);
            AppendHashValue(builder, "visual.payloadHash", visual.PayloadHash);
            AppendLocationHashInput(builder, "visual.location", visual.Location);
            AppendHashValue(builder, "visual.sourceName", visual.SourceName);
            AppendHashValue(builder, "visual.mimeType", visual.MimeType);
            AppendHashValue(builder, "visual.width", visual.Width);
            AppendHashValue(builder, "visual.height", visual.Height);
            AppendHashValue(builder, "visual.x", visual.X);
            AppendHashValue(builder, "visual.y", visual.Y);
            AppendHashValue(builder, "visual.placedWidth", visual.PlacedWidth);
            AppendHashValue(builder, "visual.placedHeight", visual.PlacedHeight);
            AppendHashValue(builder, "visual.placementCount", visual.PlacementCount);
            AppendHashValue(builder, "visual.hasGeometry", visual.HasGeometry);
            AppendHashValue(builder, "visual.isAxisAligned", visual.IsAxisAligned);
        }
    }

    private static void AppendFormFieldsHashInput(StringBuilder builder, IReadOnlyList<ReaderFormField>? fields) {
        AppendHashValue(builder, "formFields.count", fields?.Count ?? 0);
        if (fields is null) return;

        for (int i = 0; i < fields.Count; i++) {
            ReaderFormField field = fields[i];
            AppendHashValue(builder, "formField.index", i);
            AppendHashValue(builder, "formField.name", field.Name);
            AppendHashValue(builder, "formField.partialName", field.PartialName);
            AppendHashValue(builder, "formField.alternateName", field.AlternateName);
            AppendHashValue(builder, "formField.mappingName", field.MappingName);
            AppendHashValue(builder, "formField.fieldType", field.FieldType);
            AppendHashValue(builder, "formField.kind", field.Kind);
            AppendHashValue(builder, "formField.value", field.Value);
            AppendStringListHashInput(builder, "formField.values", field.Values);
            AppendHashValue(builder, "formField.defaultValue", field.DefaultValue);
            AppendStringListHashInput(builder, "formField.defaultValues", field.DefaultValues);
            AppendHashValue(builder, "formField.maxLength", field.MaxLength);
            AppendHashValue(builder, "formField.isReadOnly", field.IsReadOnly);
            AppendHashValue(builder, "formField.isRequired", field.IsRequired);
            AppendHashValue(builder, "formField.isNoExport", field.IsNoExport);
            AppendHashValue(builder, "formField.isMultiline", field.IsMultiline);
            AppendHashValue(builder, "formField.isPassword", field.IsPassword);
            AppendHashValue(builder, "formField.isComb", field.IsComb);
            AppendHashValue(builder, "formField.optionCount", field.OptionCount);
            AppendHashValue(builder, "formField.selectedOptionCount", field.SelectedOptionCount);
            AppendHashValue(builder, "formField.widgetCount", field.WidgetCount);
            AppendIntListHashInput(builder, "formField.pageNumbers", field.PageNumbers);
            AppendHashValue(builder, "formField.widgets.count", field.Widgets.Count);
            for (int widgetIndex = 0; widgetIndex < field.Widgets.Count; widgetIndex++) {
                ReaderFormWidget widget = field.Widgets[widgetIndex];
                AppendHashValue(builder, "formWidget.index", widgetIndex);
                AppendHashValue(builder, "formWidget.fieldName", widget.FieldName);
                AppendHashValue(builder, "formWidget.pageNumber", widget.PageNumber);
                AppendHashValue(builder, "formWidget.x1", widget.X1);
                AppendHashValue(builder, "formWidget.y1", widget.Y1);
                AppendHashValue(builder, "formWidget.x2", widget.X2);
                AppendHashValue(builder, "formWidget.y2", widget.Y2);
                AppendHashValue(builder, "formWidget.width", widget.Width);
                AppendHashValue(builder, "formWidget.height", widget.Height);
                AppendHashValue(builder, "formWidget.appearanceState", widget.AppearanceState);
                AppendHashValue(builder, "formWidget.isHidden", widget.IsHidden);
                AppendHashValue(builder, "formWidget.isPrint", widget.IsPrint);
                AppendHashValue(builder, "formWidget.isReadOnly", widget.IsReadOnly);
                AppendHashValue(builder, "formWidget.normalAppearanceStateCount", widget.NormalAppearanceStateCount);
                AppendStringListHashInput(builder, "formWidget.normalAppearanceStates", widget.NormalAppearanceStates);
            }
        }
    }

    private static void AppendActionsHashInput(StringBuilder builder, IReadOnlyList<ReaderActionSummary>? actions) {
        AppendHashValue(builder, "actions.count", actions?.Count ?? 0);
        if (actions is null) return;

        for (int i = 0; i < actions.Count; i++) {
            ReaderActionSummary action = actions[i];
            AppendHashValue(builder, "action.index", i);
            AppendHashValue(builder, "action.scope", action.Scope);
            AppendHashValue(builder, "action.type", action.ActionType);
            AppendHashValue(builder, "action.source", action.Source);
            AppendHashValue(builder, "action.name", action.Name);
            AppendHashValue(builder, "action.triggerName", action.TriggerName);
            AppendHashValue(builder, "action.path", action.ActionPath);
            AppendHashValue(builder, "action.pageNumber", action.PageNumber);
            AppendHashValue(builder, "action.isChained", action.IsChainedAction);
            AppendHashValue(builder, "action.isPotentiallyUnsafe", action.IsPotentiallyUnsafe);
            AppendHashValue(builder, "action.destinationPageNumber", action.DestinationPageNumber);
            AppendHashValue(builder, "action.destinationMode", action.DestinationMode);
            AppendHashValue(builder, "action.destinationTop", action.DestinationTop);
            AppendHashValue(builder, "action.destinationLeft", action.DestinationLeft);
            AppendHashValue(builder, "action.destinationBottom", action.DestinationBottom);
            AppendHashValue(builder, "action.destinationRight", action.DestinationRight);
        }
    }

    private static void AppendDiagnosticsHashInput(StringBuilder builder, ReaderChunkDiagnostics? diagnostics) {
        AppendHashValue(builder, "diagnostics.hasValue", diagnostics is not null);
        if (diagnostics is null) return;

        AppendHashValue(builder, "diagnostics.sourceKind", diagnostics.SourceKind);
        AppendHashValue(builder, "diagnostics.pageCount", diagnostics.PageCount);
        AppendHashValue(builder, "diagnostics.selectedPageCount", diagnostics.SelectedPageCount);
        AppendHashValue(builder, "diagnostics.pageNumber", diagnostics.PageNumber);
        AppendHashValue(builder, "diagnostics.tableCount", diagnostics.TableCount);
        AppendHashValue(builder, "diagnostics.tableGeometryCount", diagnostics.TableGeometryCount);
        AppendHashValue(builder, "diagnostics.tableGeometryCoverage", diagnostics.TableGeometryCoverage);
        AppendHashValue(builder, "diagnostics.minTableConfidence", diagnostics.MinTableConfidence);
        AppendHashValue(builder, "diagnostics.averageTableConfidence", diagnostics.AverageTableConfidence);
        AppendHashValue(builder, "diagnostics.lowConfidenceTableCount", diagnostics.LowConfidenceTableCount);
        AppendHashValue(builder, "diagnostics.numericTableColumnCount", diagnostics.NumericTableColumnCount);
        AppendHashValue(builder, "diagnostics.fallbackTableColumnNameCount", diagnostics.FallbackTableColumnNameCount);
        AppendHashValue(builder, "diagnostics.missingTableCellCount", diagnostics.MissingTableCellCount);
        AppendHashValue(builder, "diagnostics.imageCount", diagnostics.ImageCount);
        AppendHashValue(builder, "diagnostics.imageGeometryCount", diagnostics.ImageGeometryCount);
        AppendHashValue(builder, "diagnostics.imageGeometryCoverage", diagnostics.ImageGeometryCoverage);
        AppendHashValue(builder, "diagnostics.imageNonAxisAlignedCount", diagnostics.ImageNonAxisAlignedCount);
        AppendHashValue(builder, "diagnostics.imageNonAxisAlignedCoverage", diagnostics.ImageNonAxisAlignedCoverage);
        AppendHashValue(builder, "diagnostics.linkCount", diagnostics.LinkCount);
        AppendHashValue(builder, "diagnostics.hasXmpMetadata", diagnostics.HasXmpMetadata);
        AppendHashValue(builder, "diagnostics.outputIntentCount", diagnostics.OutputIntentCount);
        AppendHashValue(builder, "diagnostics.attachmentCount", diagnostics.AttachmentCount);
        AppendHashValue(builder, "diagnostics.hasTaggedContent", diagnostics.HasTaggedContent);
        AppendHashValue(builder, "diagnostics.taggedStructureElementCount", diagnostics.TaggedStructureElementCount);
        AppendHashValue(builder, "diagnostics.taggedMarkedContentReferenceCount", diagnostics.TaggedMarkedContentReferenceCount);
        AppendHashValue(builder, "diagnostics.optionalContentGroupCount", diagnostics.OptionalContentGroupCount);
        AppendHashValue(builder, "diagnostics.optionalContentInitiallyHiddenCount", diagnostics.OptionalContentInitiallyHiddenCount);
        AppendHashValue(builder, "diagnostics.optionalContentLockedCount", diagnostics.OptionalContentLockedCount);
        AppendHashValue(builder, "diagnostics.hasOpenAction", diagnostics.HasOpenAction);
        AppendHashValue(builder, "diagnostics.hasCatalogActions", diagnostics.HasCatalogActions);
        AppendHashValue(builder, "diagnostics.hasPageActions", diagnostics.HasPageActions);
        AppendHashValue(builder, "diagnostics.hasAnnotationActions", diagnostics.HasAnnotationActions);
        AppendHashValue(builder, "diagnostics.hasActiveContent", diagnostics.HasActiveContent);
        AppendHashValue(builder, "diagnostics.potentiallyUnsafeActionCount", diagnostics.PotentiallyUnsafeActionCount);
        AppendHashValue(builder, "diagnostics.javaScriptActionCount", diagnostics.JavaScriptActionCount);
        AppendHashValue(builder, "diagnostics.launchActionCount", diagnostics.LaunchActionCount);
        AppendHashValue(builder, "diagnostics.submitFormActionCount", diagnostics.SubmitFormActionCount);
        AppendHashValue(builder, "diagnostics.importDataActionCount", diagnostics.ImportDataActionCount);
        AppendHashValue(builder, "diagnostics.catalogActionCount", diagnostics.CatalogActionCount);
        AppendHashValue(builder, "diagnostics.pageActionCount", diagnostics.PageActionCount);
        AppendHashValue(builder, "diagnostics.selectedPageActionCount", diagnostics.SelectedPageActionCount);
        AppendHashValue(builder, "diagnostics.annotationActionCount", diagnostics.AnnotationActionCount);
        AppendHashValue(builder, "diagnostics.selectedAnnotationActionCount", diagnostics.SelectedAnnotationActionCount);
        AppendHashValue(builder, "diagnostics.formFieldCount", diagnostics.FormFieldCount);
        AppendHashValue(builder, "diagnostics.formWidgetCount", diagnostics.FormWidgetCount);
        AppendHashValue(builder, "diagnostics.selectedFormWidgetCount", diagnostics.SelectedFormWidgetCount);
        AppendHashValue(builder, "diagnostics.selectedFormWidgetAppearanceStateCount", diagnostics.SelectedFormWidgetAppearanceStateCount);
        AppendHashValue(builder, "diagnostics.selectedFormWidgetAppearanceStateCoverage", diagnostics.SelectedFormWidgetAppearanceStateCoverage);
        AppendHashValue(builder, "diagnostics.selectedFormWidgetNormalAppearanceStateCount", diagnostics.SelectedFormWidgetNormalAppearanceStateCount);
        AppendHashValue(builder, "diagnostics.hasSecurityState", diagnostics.HasSecurityState);
        AppendHashValue(builder, "diagnostics.hasEncryption", diagnostics.HasEncryption);
        AppendHashValue(builder, "diagnostics.hasSignatures", diagnostics.HasSignatures);
        AppendHashValue(builder, "diagnostics.hasIncrementalUpdates", diagnostics.HasIncrementalUpdates);
        AppendHashValue(builder, "diagnostics.revisionCount", diagnostics.RevisionCount);
        AppendHashValue(builder, "diagnostics.requiresAppendOnlyMutation", diagnostics.RequiresAppendOnlyMutation);
    }

    private static void AppendLocationHashInput(StringBuilder builder, string prefix, ReaderLocation? location) {
        AppendHashValue(builder, prefix + ".hasValue", location is not null);
        if (location is null) return;

        AppendHashValue(builder, prefix + ".path", location.Path);
        AppendHashValue(builder, prefix + ".blockIndex", location.BlockIndex);
        AppendHashValue(builder, prefix + ".sourceBlockIndex", location.SourceBlockIndex);
        AppendHashValue(builder, prefix + ".startLine", location.StartLine);
        AppendHashValue(builder, prefix + ".endLine", location.EndLine);
        AppendHashValue(builder, prefix + ".normalizedStartLine", location.NormalizedStartLine);
        AppendHashValue(builder, prefix + ".normalizedEndLine", location.NormalizedEndLine);
        AppendHashValue(builder, prefix + ".headingPath", location.HeadingPath);
        AppendHashValue(builder, prefix + ".headingSlug", location.HeadingSlug);
        AppendHashValue(builder, prefix + ".sourceBlockKind", location.SourceBlockKind);
        AppendHashValue(builder, prefix + ".blockAnchor", location.BlockAnchor);
        AppendHashValue(builder, prefix + ".sheet", location.Sheet);
        AppendHashValue(builder, prefix + ".a1Range", location.A1Range);
        AppendHashValue(builder, prefix + ".slide", location.Slide);
        AppendHashValue(builder, prefix + ".page", location.Page);
        AppendHashValue(builder, prefix + ".tableIndex", location.TableIndex);
    }

    private static void AppendStringListHashInput(StringBuilder builder, string name, IReadOnlyList<string>? values) {
        AppendHashValue(builder, name + ".count", values?.Count ?? 0);
        if (values is null) return;

        for (int i = 0; i < values.Count; i++) {
            AppendHashValue(builder, name + "." + i.ToString(CultureInfo.InvariantCulture), values[i]);
        }
    }

    private static void AppendIntListHashInput(StringBuilder builder, string name, IReadOnlyList<int>? values) {
        AppendHashValue(builder, name + ".count", values?.Count ?? 0);
        if (values is null) return;

        for (int i = 0; i < values.Count; i++) {
            AppendHashValue(builder, name + "." + i.ToString(CultureInfo.InvariantCulture), values[i]);
        }
    }

    private static void AppendHashValue(StringBuilder builder, string name, object? value) {
        builder.Append(name);
        builder.Append('=');
        if (value is IFormattable formattable) {
            builder.Append(formattable.ToString(null, CultureInfo.InvariantCulture));
        } else if (value is not null) {
            builder.Append(value);
        }

        builder.Append(';');
    }
}
