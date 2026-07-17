using System;
using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeIMO.Reader;

/// <summary>
/// JSON serialization helpers for the shared OfficeIMO document read result envelope.
/// </summary>
public static partial class OfficeDocumentReadResultJson {
    private static readonly string[] RequiredTopLevelProperties = {
        "schemaId",
        "schemaVersion",
        "kind",
        "source",
        "capabilitiesUsed",
        "chunks",
        "metadata",
        "pages",
        "blocks",
        "tables",
        "assets",
        "links",
        "forms",
        "ocrCandidates",
        "visuals",
        "diagnostics"
    };

    private static readonly HashSet<string> AllowedTopLevelProperties = new HashSet<string>(
        new[] {
            "schemaId",
            "schemaVersion",
            "kind",
            "source",
            "capabilitiesUsed",
            "markdown",
            "html",
            "json",
            "chunks",
            "metadata",
            "pages",
            "blocks",
            "tables",
            "assets",
            "links",
            "forms",
            "ocrCandidates",
            "visuals",
            "diagnostics"
        },
        StringComparer.Ordinal);

    /// <summary>
    /// Serializes a document read result into the stable OfficeIMO transport shape.
    /// </summary>
    /// <param name="result">Read result to serialize.</param>
    /// <param name="indented">When true, writes indented JSON for diagnostics and fixtures.</param>
    public static string Serialize(OfficeDocumentReadResult result, bool indented = false) {
        if (result == null) throw new ArgumentNullException(nameof(result));
        string schemaId = string.IsNullOrWhiteSpace(result.SchemaId)
            ? OfficeDocumentReadResultSchema.Id
            : result.SchemaId;
        int schemaVersion = result.SchemaVersion == 0
            ? OfficeDocumentReadResultSchema.CurrentVersion
            : result.SchemaVersion;
        OfficeDocumentReadResultSchema.EnsureSupported(schemaId, schemaVersion);
        EnsureKindSupported(schemaVersion, result.Kind);
        EnsureChunkKindsSupported(schemaVersion, result.Chunks);
        EnsureStringCollection(result.CapabilitiesUsed, "capabilitiesUsed");
        EnsureDiagnosticContracts(result.Diagnostics);

        return JsonSerializer.Serialize(ProjectResult(result), CreateOptions(indented));
    }

    /// <summary>
    /// Serializes a document read result into the stable OfficeIMO transport shape.
    /// </summary>
    /// <param name="result">Read result to serialize.</param>
    /// <param name="indented">When true, writes indented JSON for diagnostics and fixtures.</param>
    public static string ToJson(this OfficeDocumentReadResult result, bool indented = false) {
        return Serialize(result, indented);
    }

    /// <summary>
    /// Deserializes a current, supported document read result transport payload.
    /// </summary>
    /// <param name="json">UTF-16 JSON text containing one complete result envelope.</param>
    /// <exception cref="JsonException">The payload is not valid JSON or cannot be mapped to the transport model.</exception>
    /// <exception cref="OfficeDocumentReadResultSchemaException">The schema identifier or version is not supported.</exception>
    public static OfficeDocumentReadResult Deserialize(string json) {
        if (json == null) throw new ArgumentNullException(nameof(json));

        using JsonDocument document = JsonDocument.Parse(json);
        JsonElement root = document.RootElement;
        if (root.ValueKind != JsonValueKind.Object) {
            throw new JsonException("The document read result payload must be a JSON object.");
        }

        string? schemaId = TryReadSchemaId(root);
        int schemaVersion = TryReadSchemaVersion(root);
        OfficeDocumentReadResultSchema.EnsureSupported(schemaId, schemaVersion);
        EnsureRequiredTopLevelProperties(root);
        EnsureKnownTopLevelProperties(root);
        EnsureNestedTransportContracts(root);

        OfficeDocumentReadResult? result = JsonSerializer.Deserialize<OfficeDocumentReadResult>(json, CreateReadOptions());
        if (result == null) {
            throw new JsonException("The document read result payload produced a null result.");
        }
        EnsureKindSupported(schemaVersion, result.Kind);
        EnsureChunkKindsSupported(schemaVersion, result.Chunks);
        result = NormalizeDeserializedResult(result);
        EnsureDiagnosticContracts(result.Diagnostics);
        return result;
    }

    private static JsonSerializerOptions CreateOptions(bool indented) {
        return new JsonSerializerOptions {
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
            WriteIndented = indented
        };
    }

    private static JsonSerializerOptions CreateReadOptions() {
        var options = new JsonSerializerOptions {
            PropertyNameCaseInsensitive = true
        };
        options.Converters.Add(new JsonStringEnumConverter(namingPolicy: null, allowIntegerValues: false));
        return options;
    }

    private static void EnsureRequiredTopLevelProperties(JsonElement root) {
        for (int index = 0; index < RequiredTopLevelProperties.Length; index++) {
            string propertyName = RequiredTopLevelProperties[index];
            if (!root.TryGetProperty(propertyName, out JsonElement property)) {
                throw new JsonException($"Required document read result property '{propertyName}' is missing.");
            }
            if (property.ValueKind == JsonValueKind.Null) {
                throw new JsonException($"Required document read result property '{propertyName}' cannot be null.");
            }
        }
    }

    private static void EnsureKnownTopLevelProperties(JsonElement root) {
        foreach (JsonProperty property in root.EnumerateObject()) {
            if (!AllowedTopLevelProperties.Contains(property.Name)) {
                throw new JsonException($"Unknown document read result property '{property.Name}'.");
            }
        }
    }

    private static void EnsureDiagnosticContracts(IReadOnlyList<OfficeDocumentDiagnostic>? diagnostics) {
        if (diagnostics == null) return;

        for (int index = 0; index < diagnostics.Count; index++) {
            OfficeDocumentDiagnostic? diagnostic = diagnostics[index];
            if (diagnostic == null || string.IsNullOrWhiteSpace(diagnostic.Code)) {
                throw new JsonException($"Document diagnostic at index {index} must have a non-empty code.");
            }
            if (diagnostic.Message == null) {
                throw new JsonException($"Document diagnostic at index {index} must have a message string.");
            }
            if (diagnostic.Attributes == null) {
                throw new JsonException($"Document diagnostic at index {index} must have an attributes object.");
            }
            foreach (KeyValuePair<string, string> attribute in diagnostic.Attributes) {
                if (attribute.Value == null) {
                    throw new JsonException($"Document diagnostic at index {index} has a null attribute value for '{attribute.Key}'.");
                }
            }
        }
    }

    private static void EnsureKindSupported(int schemaVersion, ReaderInputKind kind) {
        if (!Enum.IsDefined(typeof(ReaderInputKind), kind) ||
            schemaVersion == 5 && (kind == ReaderInputKind.Calendar || kind == ReaderInputKind.VCard)) {
            throw new JsonException(
                $"Reader input kind '{kind}' is not supported by document read result schema version {schemaVersion}.");
        }
    }

    private static void EnsureChunkKindsSupported(int schemaVersion, IReadOnlyList<ReaderChunk>? chunks) {
        if (chunks == null) return;
        for (int index = 0; index < chunks.Count; index++) {
            ReaderChunk? chunk = chunks[index];
            if (chunk != null) EnsureKindSupported(schemaVersion, chunk.Kind);
        }
    }

    private static string? TryReadSchemaId(JsonElement root) {
        if (!root.TryGetProperty("schemaId", out JsonElement property) ||
            property.ValueKind != JsonValueKind.String) {
            return null;
        }
        return property.GetString();
    }

    private static int TryReadSchemaVersion(JsonElement root) {
        if (!root.TryGetProperty("schemaVersion", out JsonElement property) ||
            property.ValueKind != JsonValueKind.Number ||
            !property.TryGetInt32(out int version)) {
            return 0;
        }
        return version;
    }

    private static OfficeDocumentReadResult NormalizeDeserializedResult(OfficeDocumentReadResult result) {
        result.SchemaId = OfficeDocumentReadResultSchema.Id;
        result.SchemaVersion = OfficeDocumentReadResultSchema.CurrentVersion;
        result.Source ??= new OfficeDocumentSource();
        result.CapabilitiesUsed ??= Array.Empty<string>();
        result.Chunks ??= Array.Empty<ReaderChunk>();
        result.Metadata ??= Array.Empty<OfficeDocumentMetadataEntry>();
        result.Pages ??= Array.Empty<OfficeDocumentPage>();
        result.Blocks ??= Array.Empty<OfficeDocumentBlock>();
        result.Tables ??= Array.Empty<ReaderTable>();
        result.Assets ??= Array.Empty<OfficeDocumentAsset>();
        result.Links ??= Array.Empty<OfficeDocumentLink>();
        result.Forms ??= Array.Empty<OfficeDocumentFormField>();
        result.OcrCandidates ??= Array.Empty<OfficeDocumentOcrCandidate>();
        result.Visuals ??= Array.Empty<ReaderVisual>();
        result.Diagnostics ??= Array.Empty<OfficeDocumentDiagnostic>();
        return result;
    }

    private static object ProjectResult(OfficeDocumentReadResult result) {
        return new {
            schemaId = string.IsNullOrWhiteSpace(result.SchemaId) ? OfficeDocumentReadResultSchema.Id : result.SchemaId,
            schemaVersion = result.SchemaVersion == 0 ? OfficeDocumentReadResultSchema.CurrentVersion : result.SchemaVersion,
            kind = result.Kind.ToString(),
            source = ProjectSource(result.Source ?? new OfficeDocumentSource()),
            capabilitiesUsed = OrEmpty(result.CapabilitiesUsed),
            markdown = result.Markdown,
            html = result.Html,
            json = result.Json,
            chunks = ProjectCollection(result.Chunks, ProjectChunk),
            metadata = ProjectCollection(result.Metadata, ProjectMetadataEntry),
            pages = ProjectCollection(result.Pages, ProjectPage),
            blocks = ProjectCollection(result.Blocks, ProjectBlock),
            tables = ProjectCollection(result.Tables, ProjectTable),
            assets = ProjectCollection(result.Assets, ProjectAsset),
            links = ProjectCollection(result.Links, ProjectLink),
            forms = ProjectCollection(result.Forms, ProjectForm),
            ocrCandidates = ProjectCollection(result.OcrCandidates, ProjectOcrCandidate),
            visuals = ProjectCollection(result.Visuals, ProjectVisual),
            diagnostics = ProjectCollection(result.Diagnostics, ProjectDiagnostic)
        };
    }

    private static object ProjectMetadataEntry(OfficeDocumentMetadataEntry entry) {
        return new {
            id = entry.Id,
            category = entry.Category,
            name = entry.Name,
            value = entry.Value,
            valueType = entry.ValueType,
            sourceObjectId = entry.SourceObjectId,
            location = ProjectLocation(entry.Location),
            attributes = ProjectAttributes(entry.Attributes)
        };
    }

    private static object ProjectSource(OfficeDocumentSource source) {
        return new {
            path = source.Path,
            sourceId = source.SourceId,
            sourceHash = source.SourceHash,
            lastWriteUtc = source.LastWriteUtc,
            lengthBytes = source.LengthBytes,
            title = source.Title,
            author = source.Author,
            subject = source.Subject,
            keywords = source.Keywords
        };
    }

    private static object ProjectChunk(ReaderChunk chunk) {
        return new {
            id = chunk.Id,
            kind = chunk.Kind.ToString(),
            location = ProjectLocation(chunk.Location),
            sourceId = chunk.SourceId,
            sourceHash = chunk.SourceHash,
            chunkHash = chunk.ChunkHash,
            sourceLastWriteUtc = chunk.SourceLastWriteUtc,
            sourceLengthBytes = chunk.SourceLengthBytes,
            tokenEstimate = chunk.TokenEstimate,
            text = chunk.Text,
            markdown = chunk.Markdown,
            tables = ProjectNullableCollection(chunk.Tables, ProjectTable),
            visuals = ProjectNullableCollection(chunk.Visuals, ProjectVisual),
            formFields = ProjectNullableCollection(chunk.FormFields, ProjectFormField),
            actions = ProjectNullableCollection(chunk.Actions, ProjectAction),
            diagnostics = ProjectChunkDiagnostics(chunk.Diagnostics),
            warnings = chunk.Warnings
        };
    }

    private static object ProjectPage(OfficeDocumentPage page) {
        return new {
            number = page.Number,
            name = page.Name,
            width = page.Width,
            height = page.Height,
            rotationDegrees = page.RotationDegrees,
            location = ProjectLocation(page.Location),
            blocks = ProjectCollection(page.Blocks, ProjectBlock),
            tables = ProjectCollection(page.Tables, ProjectTable),
            assets = ProjectCollection(page.Assets, ProjectAsset),
            links = ProjectCollection(page.Links, ProjectLink),
            forms = ProjectCollection(page.Forms, ProjectForm),
            ocrCandidates = ProjectCollection(page.OcrCandidates, ProjectOcrCandidate)
        };
    }

    private static object ProjectBlock(OfficeDocumentBlock block) {
        return new {
            id = block.Id,
            kind = block.Kind,
            text = block.Text,
            level = block.Level,
            marker = block.Marker,
            location = ProjectLocation(block.Location),
            region = ProjectRegion(block.Region)
        };
    }

    private static object ProjectTable(ReaderTable table) {
        return new {
            title = table.Title,
            kind = table.Kind,
            callId = table.CallId,
            summary = table.Summary,
            payloadHash = table.PayloadHash,
            location = ProjectLocation(table.Location),
            columns = OrEmpty(table.Columns),
            columnProfiles = ProjectCollection(table.ColumnProfiles, ProjectColumnProfile),
            diagnostics = ProjectTableDiagnostics(table.Diagnostics),
            rows = ProjectRows(table.Rows),
            totalRowCount = table.TotalRowCount,
            truncated = table.Truncated
        };
    }

    private static object ProjectColumnProfile(ReaderTableColumnProfile profile) {
        return new {
            index = profile.Index,
            name = profile.Name,
            kind = profile.Kind.ToString(),
            nonEmptyCellCount = profile.NonEmptyCellCount,
            numericCellCount = profile.NumericCellCount,
            confidence = profile.Confidence,
            isNumeric = profile.IsNumeric
        };
    }

    private static object ProjectAsset(OfficeDocumentAsset asset) {
        return new {
            id = asset.Id,
            kind = asset.Kind,
            mediaType = asset.MediaType,
            extension = asset.Extension,
            fileName = asset.FileName,
            altText = asset.AltText,
            title = asset.Title,
            width = asset.Width,
            height = asset.Height,
            lengthBytes = asset.LengthBytes,
            payloadHash = asset.PayloadHash,
            sourceObjectId = asset.SourceObjectId,
            location = ProjectLocation(asset.Location),
            region = ProjectRegion(asset.Region)
        };
    }

    private static object ProjectLink(OfficeDocumentLink link) {
        return new {
            id = link.Id,
            kind = link.Kind,
            uri = link.Uri,
            destinationName = link.DestinationName,
            destinationPageNumber = link.DestinationPageNumber,
            destinationMode = link.DestinationMode,
            destinationTop = link.DestinationTop,
            destinationLeft = link.DestinationLeft,
            destinationBottom = link.DestinationBottom,
            destinationRight = link.DestinationRight,
            namedAction = link.NamedAction,
            remoteFile = link.RemoteFile,
            remoteDestinationName = link.RemoteDestinationName,
            remoteDestinationPageNumber = link.RemoteDestinationPageNumber,
            remoteDestinationMode = link.RemoteDestinationMode,
            remoteDestinationTop = link.RemoteDestinationTop,
            remoteDestinationLeft = link.RemoteDestinationLeft,
            remoteDestinationBottom = link.RemoteDestinationBottom,
            remoteDestinationRight = link.RemoteDestinationRight,
            text = link.Text,
            location = ProjectLocation(link.Location),
            region = ProjectRegion(link.Region)
        };
    }

    private static object ProjectForm(OfficeDocumentFormField form) {
        return new {
            id = form.Id,
            name = form.Name,
            kind = form.Kind,
            value = form.Value,
            isReadOnly = form.IsReadOnly,
            isRequired = form.IsRequired,
            location = ProjectLocation(form.Location),
            region = ProjectRegion(form.Region)
        };
    }

    private static object ProjectVisual(ReaderVisual visual) {
        return new {
            kind = visual.Kind,
            language = visual.Language,
            content = visual.Content,
            payloadHash = visual.PayloadHash,
            sourceName = visual.SourceName,
            mimeType = visual.MimeType,
            width = visual.Width,
            height = visual.Height,
            x = visual.X,
            y = visual.Y,
            placedWidth = visual.PlacedWidth,
            placedHeight = visual.PlacedHeight,
            placementCount = visual.PlacementCount,
            hasGeometry = visual.HasGeometry,
            isAxisAligned = visual.IsAxisAligned,
            location = ProjectLocation(visual.Location)
        };
    }

    private static object ProjectFormField(ReaderFormField field) {
        return new {
            name = field.Name,
            partialName = field.PartialName,
            alternateName = field.AlternateName,
            mappingName = field.MappingName,
            fieldType = field.FieldType,
            kind = field.Kind.ToString(),
            value = field.Value,
            values = OrEmpty(field.Values),
            defaultValue = field.DefaultValue,
            defaultValues = OrEmpty(field.DefaultValues),
            maxLength = field.MaxLength,
            isReadOnly = field.IsReadOnly,
            isRequired = field.IsRequired,
            isNoExport = field.IsNoExport,
            isMultiline = field.IsMultiline,
            isPassword = field.IsPassword,
            isComb = field.IsComb,
            optionCount = field.OptionCount,
            selectedOptionCount = field.SelectedOptionCount,
            widgetCount = field.WidgetCount,
            pageNumbers = OrEmpty(field.PageNumbers),
            widgets = ProjectCollection(field.Widgets, ProjectFormWidget)
        };
    }

    private static object ProjectFormWidget(ReaderFormWidget widget) {
        return new {
            fieldName = widget.FieldName,
            pageNumber = widget.PageNumber,
            x1 = widget.X1,
            y1 = widget.Y1,
            x2 = widget.X2,
            y2 = widget.Y2,
            width = widget.Width,
            height = widget.Height,
            appearanceState = widget.AppearanceState,
            isHidden = widget.IsHidden,
            isPrint = widget.IsPrint,
            isReadOnly = widget.IsReadOnly,
            normalAppearanceStateCount = widget.NormalAppearanceStateCount,
            normalAppearanceStates = OrEmpty(widget.NormalAppearanceStates)
        };
    }

    private static object ProjectAction(ReaderActionSummary action) {
        return new {
            scope = action.Scope.ToString(),
            actionType = action.ActionType,
            source = action.Source,
            name = action.Name,
            triggerName = action.TriggerName,
            actionPath = action.ActionPath,
            pageNumber = action.PageNumber,
            isChainedAction = action.IsChainedAction,
            isPotentiallyUnsafe = action.IsPotentiallyUnsafe,
            destinationPageNumber = action.DestinationPageNumber,
            destinationMode = action.DestinationMode,
            destinationTop = action.DestinationTop,
            destinationLeft = action.DestinationLeft,
            destinationBottom = action.DestinationBottom,
            destinationRight = action.DestinationRight
        };
    }

    private static object? ProjectChunkDiagnostics(ReaderChunkDiagnostics? diagnostics) {
        if (diagnostics == null) return null;

        return new {
            sourceKind = diagnostics.SourceKind,
            pageCount = diagnostics.PageCount,
            selectedPageCount = diagnostics.SelectedPageCount,
            pageNumber = diagnostics.PageNumber,
            tableCount = diagnostics.TableCount,
            tableGeometryCount = diagnostics.TableGeometryCount,
            tableGeometryCoverage = diagnostics.TableGeometryCoverage,
            minTableConfidence = diagnostics.MinTableConfidence,
            averageTableConfidence = diagnostics.AverageTableConfidence,
            lowConfidenceTableCount = diagnostics.LowConfidenceTableCount,
            numericTableColumnCount = diagnostics.NumericTableColumnCount,
            fallbackTableColumnNameCount = diagnostics.FallbackTableColumnNameCount,
            missingTableCellCount = diagnostics.MissingTableCellCount,
            imageCount = diagnostics.ImageCount,
            imageGeometryCount = diagnostics.ImageGeometryCount,
            imageGeometryCoverage = diagnostics.ImageGeometryCoverage,
            imageNonAxisAlignedCount = diagnostics.ImageNonAxisAlignedCount,
            imageNonAxisAlignedCoverage = diagnostics.ImageNonAxisAlignedCoverage,
            linkCount = diagnostics.LinkCount,
            hasXmpMetadata = diagnostics.HasXmpMetadata,
            outputIntentCount = diagnostics.OutputIntentCount,
            attachmentCount = diagnostics.AttachmentCount,
            hasTaggedContent = diagnostics.HasTaggedContent,
            taggedStructureElementCount = diagnostics.TaggedStructureElementCount,
            taggedMarkedContentReferenceCount = diagnostics.TaggedMarkedContentReferenceCount,
            optionalContentGroupCount = diagnostics.OptionalContentGroupCount,
            optionalContentInitiallyHiddenCount = diagnostics.OptionalContentInitiallyHiddenCount,
            optionalContentLockedCount = diagnostics.OptionalContentLockedCount,
            hasOpenAction = diagnostics.HasOpenAction,
            hasCatalogActions = diagnostics.HasCatalogActions,
            hasPageActions = diagnostics.HasPageActions,
            hasAnnotationActions = diagnostics.HasAnnotationActions,
            hasActiveContent = diagnostics.HasActiveContent,
            potentiallyUnsafeActionCount = diagnostics.PotentiallyUnsafeActionCount,
            javaScriptActionCount = diagnostics.JavaScriptActionCount,
            launchActionCount = diagnostics.LaunchActionCount,
            submitFormActionCount = diagnostics.SubmitFormActionCount,
            importDataActionCount = diagnostics.ImportDataActionCount,
            catalogActionCount = diagnostics.CatalogActionCount,
            pageActionCount = diagnostics.PageActionCount,
            selectedPageActionCount = diagnostics.SelectedPageActionCount,
            annotationActionCount = diagnostics.AnnotationActionCount,
            selectedAnnotationActionCount = diagnostics.SelectedAnnotationActionCount,
            formFieldCount = diagnostics.FormFieldCount,
            formWidgetCount = diagnostics.FormWidgetCount,
            selectedFormWidgetCount = diagnostics.SelectedFormWidgetCount,
            selectedFormWidgetAppearanceStateCount = diagnostics.SelectedFormWidgetAppearanceStateCount,
            selectedFormWidgetAppearanceStateCoverage = diagnostics.SelectedFormWidgetAppearanceStateCoverage,
            selectedFormWidgetNormalAppearanceStateCount = diagnostics.SelectedFormWidgetNormalAppearanceStateCount,
            hasSecurityState = diagnostics.HasSecurityState,
            hasEncryption = diagnostics.HasEncryption,
            hasSignatures = diagnostics.HasSignatures,
            hasIncrementalUpdates = diagnostics.HasIncrementalUpdates,
            revisionCount = diagnostics.RevisionCount,
            requiresAppendOnlyMutation = diagnostics.RequiresAppendOnlyMutation
        };
    }

    private static object? ProjectTableDiagnostics(ReaderTableDiagnostics? diagnostics) {
        if (diagnostics == null) return null;

        return new {
            confidence = diagnostics.Confidence,
            schemaConfidence = diagnostics.SchemaConfidence,
            cellCompleteness = diagnostics.CellCompleteness,
            columnGeometryConfidence = diagnostics.ColumnGeometryConfidence,
            sourceRowCount = diagnostics.SourceRowCount,
            expectedCellCount = diagnostics.ExpectedCellCount,
            filledCellCount = diagnostics.FilledCellCount,
            missingCellCount = diagnostics.MissingCellCount,
            xStart = diagnostics.XStart,
            xEnd = diagnostics.XEnd,
            yTop = diagnostics.YTop,
            yBottom = diagnostics.YBottom,
            width = diagnostics.Width,
            height = diagnostics.Height,
            hasGeometry = diagnostics.HasGeometry
        };
    }

    private static object ProjectOcrCandidate(OfficeDocumentOcrCandidate candidate) {
        return new {
            id = candidate.Id,
            kind = candidate.Kind,
            reason = candidate.Reason,
            confidence = candidate.Confidence,
            assetId = candidate.AssetId,
            imageCount = candidate.ImageCount,
            textBlockCount = candidate.TextBlockCount,
            location = ProjectLocation(candidate.Location),
            region = ProjectRegion(candidate.Region)
        };
    }

    private static object ProjectDiagnostic(OfficeDocumentDiagnostic diagnostic) {
        return new {
            severity = diagnostic.Severity.ToString(),
            category = diagnostic.Category.ToString(),
            code = diagnostic.Code,
            message = diagnostic.Message,
            source = diagnostic.Source,
            isRecoverable = diagnostic.IsRecoverable,
            location = ProjectLocation(diagnostic.Location),
            attributes = ProjectAttributes(diagnostic.Attributes)
        };
    }

    private static object? ProjectLocation(ReaderLocation? location) {
        if (location == null) return null;

        return new {
            path = location.Path,
            blockIndex = location.BlockIndex,
            sourceBlockIndex = location.SourceBlockIndex,
            startLine = location.StartLine,
            endLine = location.EndLine,
            normalizedStartLine = location.NormalizedStartLine,
            normalizedEndLine = location.NormalizedEndLine,
            headingPath = location.HeadingPath,
            hierarchyHeadingPath = location.HierarchyHeadingPath,
            headingSlug = location.HeadingSlug,
            sourceBlockKind = location.SourceBlockKind,
            blockAnchor = location.BlockAnchor,
            sheet = location.Sheet,
            a1Range = location.A1Range,
            slide = location.Slide,
            page = location.Page,
            tableIndex = location.TableIndex
        };
    }

    private static object? ProjectRegion(OfficeDocumentRegion? region) {
        if (region == null) return null;

        return new {
            x = region.X,
            y = region.Y,
            width = region.Width,
            height = region.Height
        };
    }

    private static IReadOnlyList<IReadOnlyList<string>> ProjectRows(IReadOnlyList<IReadOnlyList<string>>? rows) {
        return rows ?? Array.Empty<IReadOnlyList<string>>();
    }

    private static IReadOnlyDictionary<string, string> ProjectAttributes(IReadOnlyDictionary<string, string>? attributes) {
        var sorted = new SortedDictionary<string, string>(StringComparer.Ordinal);
        if (attributes == null || attributes.Count == 0) {
            return sorted;
        }

        foreach (KeyValuePair<string, string> attribute in attributes) {
            sorted[attribute.Key] = attribute.Value;
        }

        return sorted;
    }

    private static IReadOnlyList<object> ProjectCollection<T>(IReadOnlyList<T>? values, Func<T, object> projector) {
        if (values == null || values.Count == 0) {
            return Array.Empty<object>();
        }

        var projected = new object[values.Count];
        for (int i = 0; i < values.Count; i++) {
            if (ReferenceEquals(values[i], null)) {
                throw new JsonException($"Document transport collection contains a null item at index {i}.");
            }
            projected[i] = projector(values[i]);
        }

        return projected;
    }

    private static IReadOnlyList<object>? ProjectNullableCollection<T>(IReadOnlyList<T>? values, Func<T, object> projector) {
        return values == null ? null : ProjectCollection(values, projector);
    }

    private static IReadOnlyList<T> OrEmpty<T>(IReadOnlyList<T>? values) {
        return values ?? Array.Empty<T>();
    }
}
