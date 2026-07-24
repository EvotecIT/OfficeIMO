using OfficeIMO.Pdf;

namespace OfficeIMO.Reader.Pdf;

internal static partial class PdfReaderAdapter {
    private static void AddMetadata(List<OfficeDocumentMetadataEntry> entries, string id, string category, string name, string? value, string valueType = "string") {
        if (string.IsNullOrWhiteSpace(value)) {
            return;
        }

        entries.Add(new OfficeDocumentMetadataEntry {
            Id = id,
            Category = category,
            Name = name,
            Value = value,
            ValueType = valueType
        });
    }

    private static void AddCountMetadata(List<OfficeDocumentMetadataEntry> entries, string id, string category, string name, int count) {
        if (count == 0) {
            return;
        }

        entries.Add(new OfficeDocumentMetadataEntry {
            Id = id,
            Category = category,
            Name = name,
            Value = count.ToString(CultureInfo.InvariantCulture),
            ValueType = "count"
        });
    }

    private static void AddNumberMetadata(List<OfficeDocumentMetadataEntry> entries, string id, string category, string name, double value) {
        entries.Add(new OfficeDocumentMetadataEntry {
            Id = id,
            Category = category,
            Name = name,
            Value = value.ToString("R", CultureInfo.InvariantCulture),
            ValueType = "number"
        });
    }

    private static void AddImageGeometryMetadata(List<OfficeDocumentMetadataEntry> entries, IReadOnlyList<PdfLogicalPage> pages) {
        int imageCount = 0;
        int imageGeometryCount = 0;
        int nonAxisAlignedCount = 0;
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            IReadOnlyList<PdfLogicalImage> images = pages[pageIndex].Images;
            imageCount += images.Count;
            for (int imageIndex = 0; imageIndex < images.Count; imageIndex++) {
                PdfImagePlacement? placement = images[imageIndex].PrimaryPlacement;
                if (placement is not null) {
                    imageGeometryCount++;
                    if (!placement.IsAxisAligned) {
                        nonAxisAlignedCount++;
                    }
                }
            }
        }

        AddCountMetadata(entries, "pdf-image-count", "pdf.image", "Count", imageCount);
        AddCountMetadata(entries, "pdf-image-geometry-count", "pdf.image", "GeometryCount", imageGeometryCount);
        AddCountMetadata(entries, "pdf-image-non-axis-aligned-count", "pdf.image", "NonAxisAlignedCount", nonAxisAlignedCount);
        if (imageCount > 0) {
            AddNumberMetadata(entries, "pdf-image-geometry-coverage", "pdf.image", "GeometryCoverage", (double)imageGeometryCount / imageCount);
        }
        if (imageGeometryCount > 0) {
            AddNumberMetadata(entries, "pdf-image-non-axis-aligned-coverage", "pdf.image", "NonAxisAlignedCoverage", (double)nonAxisAlignedCount / imageGeometryCount);
        }
    }

    private static void AddLinkMetadata(List<OfficeDocumentMetadataEntry> entries, IReadOnlyList<PdfLogicalPage> pages) {
        int linkCount = 0;
        int linkGeometryCount = 0;
        var linkKindCounts = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            IReadOnlyList<PdfLogicalLinkAnnotation> links = pages[pageIndex].Links;
            linkCount += links.Count;
            for (int linkIndex = 0; linkIndex < links.Count; linkIndex++) {
                PdfLogicalLinkAnnotation link = links[linkIndex];
                if (link.Width > 0D && link.Height > 0D) {
                    linkGeometryCount++;
                }

                string kind = GetLinkKind(link);
                linkKindCounts[kind] = linkKindCounts.TryGetValue(kind, out int count) ? count + 1 : 1;
            }
        }

        AddCountMetadata(entries, "pdf-link-count", "pdf.link", "Count", linkCount);
        AddCountMetadata(entries, "pdf-link-geometry-count", "pdf.link", "GeometryCount", linkGeometryCount);
        if (linkCount > 0) {
            AddNumberMetadata(entries, "pdf-link-geometry-coverage", "pdf.link", "GeometryCoverage", (double)linkGeometryCount / linkCount);
        }

        foreach (KeyValuePair<string, int> linkKindCount in linkKindCounts.OrderBy(item => item.Key, StringComparer.Ordinal)) {
            AddCountMetadata(
                entries,
                "pdf-link-" + linkKindCount.Key + "-count",
                "pdf.link",
                ToMetadataDisplayName(linkKindCount.Key) + "Count",
                linkKindCount.Value);
        }
    }

    private static void AddAcroFormXfaMetadata(List<OfficeDocumentMetadataEntry> entries, PdfAcroFormXfaInfo? xfa) {
        if (xfa is null) {
            return;
        }

        AddMetadata(entries, "pdf-acroform-xfa-present", "pdf.form.xfa", "Present", "true", "boolean");
        entries.Add(new OfficeDocumentMetadataEntry {
            Id = "pdf-acroform-xfa",
            Category = "pdf.form.xfa",
            Name = "XFA",
            Value = xfa.ObjectKind,
            ValueType = "object",
            Attributes = new Dictionary<string, string>(StringComparer.Ordinal) {
                ["objectKind"] = xfa.ObjectKind,
                ["objectNumber"] = ToMetadataText(xfa.ObjectNumber) ?? string.Empty,
                ["packetCount"] = xfa.PacketCount.ToString(CultureInfo.InvariantCulture),
                ["packetNames"] = FormatPdfStringComponents(xfa.PacketNames) ?? string.Empty,
                ["streamCount"] = xfa.StreamCount.ToString(CultureInfo.InvariantCulture),
                ["stringCount"] = xfa.StringCount.ToString(CultureInfo.InvariantCulture),
                ["dictionaryCount"] = xfa.DictionaryCount.ToString(CultureInfo.InvariantCulture),
                ["totalPayloadBytes"] = xfa.TotalPayloadBytes.ToString(CultureInfo.InvariantCulture),
                ["hasTemplatePacket"] = ToMetadataText(xfa.HasTemplatePacket) ?? "false",
                ["hasDatasetsPacket"] = ToMetadataText(xfa.HasDatasetsPacket) ?? "false"
            }
        });
        AddCountMetadata(entries, "pdf-acroform-xfa-packet-count", "pdf.form.xfa", "PacketCount", xfa.PacketCount);
        AddCountMetadata(entries, "pdf-acroform-xfa-stream-count", "pdf.form.xfa", "StreamCount", xfa.StreamCount);
        AddCountMetadata(entries, "pdf-acroform-xfa-string-count", "pdf.form.xfa", "StringCount", xfa.StringCount);
        AddCountMetadata(entries, "pdf-acroform-xfa-dictionary-count", "pdf.form.xfa", "DictionaryCount", xfa.DictionaryCount);
        AddCountMetadata(entries, "pdf-acroform-xfa-payload-byte-count", "pdf.form.xfa", "PayloadByteCount", xfa.TotalPayloadBytes);
    }

    private static void AddAttachmentMetadata(List<OfficeDocumentMetadataEntry> entries, IReadOnlyList<PdfAttachmentInfo> attachments) {
        if (attachments.Count == 0) {
            return;
        }

        int associatedCount = 0;
        long totalSizeBytes = 0;
        var relationshipCounts = new Dictionary<string, int>(StringComparer.Ordinal);
        var sourceCounts = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int i = 0; i < attachments.Count; i++) {
            PdfAttachmentInfo attachment = attachments[i];
            if (attachment.IsAssociatedFile) {
                associatedCount++;
            }

            totalSizeBytes += attachment.SizeBytes;

            string relationshipKey = GetActionTypeKey(attachment.Relationship.ToString());
            relationshipCounts[relationshipKey] = relationshipCounts.TryGetValue(relationshipKey, out int relationshipCount) ? relationshipCount + 1 : 1;

            string sourceKey = GetActionTypeKey(attachment.Source);
            sourceCounts[sourceKey] = sourceCounts.TryGetValue(sourceKey, out int sourceCount) ? sourceCount + 1 : 1;
        }

        AddCountMetadata(entries, "pdf-attachment-count", "pdf.attachment", "Count", attachments.Count);
        AddCountMetadata(entries, "pdf-attachment-associated-count", "pdf.attachment", "AssociatedCount", associatedCount);
        AddMetadata(
            entries,
            "pdf-attachment-total-size-bytes",
            "pdf.attachment",
            "TotalSizeBytes",
            totalSizeBytes.ToString(CultureInfo.InvariantCulture),
            "number");

        foreach (KeyValuePair<string, int> relationshipCount in relationshipCounts.OrderBy(item => item.Key, StringComparer.Ordinal)) {
            AddCountMetadata(
                entries,
                "pdf-attachment-relationship-" + relationshipCount.Key + "-count",
                "pdf.attachment",
                ToMetadataDisplayName(relationshipCount.Key) + "RelationshipCount",
                relationshipCount.Value);
        }

        foreach (KeyValuePair<string, int> sourceCount in sourceCounts.OrderBy(item => item.Key, StringComparer.Ordinal)) {
            AddCountMetadata(
                entries,
                "pdf-attachment-source-" + sourceCount.Key + "-count",
                "pdf.attachment",
                ToMetadataDisplayName(sourceCount.Key) + "SourceCount",
                sourceCount.Value);
        }

        for (int i = 0; i < attachments.Count; i++) {
            entries.Add(BuildAttachmentMetadataEntry(attachments[i], i));
        }
    }

    private static OfficeDocumentMetadataEntry BuildAttachmentMetadataEntry(PdfAttachmentInfo attachment, int attachmentIndex) {
        string id = "pdf-attachment-" + attachmentIndex.ToString("D4", CultureInfo.InvariantCulture);
        var attributes = new Dictionary<string, string>(StringComparer.Ordinal) {
            ["name"] = attachment.Name,
            ["fileName"] = attachment.FileName,
            ["relationship"] = attachment.Relationship.ToString(),
            ["source"] = attachment.Source,
            ["isAssociatedFile"] = ToMetadataText(attachment.IsAssociatedFile),
            ["sizeBytes"] = attachment.SizeBytes.ToString(CultureInfo.InvariantCulture)
        };

        AddAttribute(attributes, "unicodeFileName", attachment.UnicodeFileName);
        AddAttribute(attributes, "description", attachment.Description);
        AddAttribute(attributes, "mimeType", attachment.MimeType);
        AddAttribute(attributes, "filter", attachment.Filter);
        AddAttribute(attributes, "fileSpecObjectNumber", attachment.FileSpecObjectNumber == 0 ? null : attachment.FileSpecObjectNumber);
        AddAttribute(attributes, "embeddedFileObjectNumber", attachment.EmbeddedFileObjectNumber == 0 ? null : attachment.EmbeddedFileObjectNumber);

        return new OfficeDocumentMetadataEntry {
            Id = id,
            Category = "pdf.attachment",
            Name = attachment.FileName,
            Value = attachment.Name,
            ValueType = "object",
            SourceObjectId = attachment.FileSpecObjectNumber == 0
                ? null
                : attachment.FileSpecObjectNumber.ToString(CultureInfo.InvariantCulture),
            Attributes = attributes
        };
    }

    private static void AddOutputIntentMetadata(List<OfficeDocumentMetadataEntry> entries, IReadOnlyList<PdfOutputIntentInfo> outputIntents) {
        if (outputIntents.Count == 0) {
            return;
        }

        int profileCount = 0;
        int iccSignatureCount = 0;
        var subtypeCounts = new Dictionary<string, int>(StringComparer.Ordinal);
        var colorSpaceCounts = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int i = 0; i < outputIntents.Count; i++) {
            PdfOutputIntentInfo outputIntent = outputIntents[i];
            if (outputIntent.HasDestinationOutputProfile) {
                profileCount++;
            }

            if (outputIntent.DestinationOutputProfileHasIccSignature == true) {
                iccSignatureCount++;
            }

            if (!string.IsNullOrWhiteSpace(outputIntent.Subtype)) {
                string subtypeKey = GetActionTypeKey(outputIntent.Subtype!);
                subtypeCounts[subtypeKey] = subtypeCounts.TryGetValue(subtypeKey, out int subtypeCount) ? subtypeCount + 1 : 1;
            }

            if (!string.IsNullOrWhiteSpace(outputIntent.DestinationOutputProfileColorSpace)) {
                string colorSpaceKey = GetActionTypeKey(outputIntent.DestinationOutputProfileColorSpace!.Trim());
                colorSpaceCounts[colorSpaceKey] = colorSpaceCounts.TryGetValue(colorSpaceKey, out int colorSpaceCount) ? colorSpaceCount + 1 : 1;
            }
        }

        AddCountMetadata(entries, "pdf-output-intent-count", "pdf.outputIntent", "Count", outputIntents.Count);
        AddCountMetadata(entries, "pdf-output-intent-profile-count", "pdf.outputIntent", "DestinationOutputProfileCount", profileCount);
        AddCountMetadata(entries, "pdf-output-intent-icc-signature-count", "pdf.outputIntent", "IccSignatureCount", iccSignatureCount);

        foreach (KeyValuePair<string, int> subtypeCount in subtypeCounts.OrderBy(item => item.Key, StringComparer.Ordinal)) {
            AddCountMetadata(
                entries,
                "pdf-output-intent-subtype-" + subtypeCount.Key + "-count",
                "pdf.outputIntent",
                ToMetadataDisplayName(subtypeCount.Key) + "SubtypeCount",
                subtypeCount.Value);
        }

        foreach (KeyValuePair<string, int> colorSpaceCount in colorSpaceCounts.OrderBy(item => item.Key, StringComparer.Ordinal)) {
            AddCountMetadata(
                entries,
                "pdf-output-intent-profile-color-space-" + colorSpaceCount.Key + "-count",
                "pdf.outputIntent",
                ToMetadataDisplayName(colorSpaceCount.Key) + "ProfileColorSpaceCount",
                colorSpaceCount.Value);
        }

        for (int i = 0; i < outputIntents.Count; i++) {
            entries.Add(BuildOutputIntentMetadataEntry(outputIntents[i], i));
        }
    }

    private static OfficeDocumentMetadataEntry BuildOutputIntentMetadataEntry(PdfOutputIntentInfo outputIntent, int outputIntentIndex) {
        string id = "pdf-output-intent-" + outputIntentIndex.ToString("D4", CultureInfo.InvariantCulture);
        var attributes = new Dictionary<string, string>(StringComparer.Ordinal);
        AddAttribute(attributes, "objectNumber", outputIntent.ObjectNumber);
        AddAttribute(attributes, "subtype", outputIntent.Subtype);
        AddAttribute(attributes, "outputConditionIdentifier", outputIntent.OutputConditionIdentifier);
        AddAttribute(attributes, "outputCondition", outputIntent.OutputCondition);
        AddAttribute(attributes, "registryName", outputIntent.RegistryName);
        AddAttribute(attributes, "info", outputIntent.Info);
        AddAttribute(attributes, "destinationOutputProfileObjectNumber", outputIntent.DestinationOutputProfileObjectNumber);
        AddAttribute(attributes, "destinationOutputProfileColorComponents", outputIntent.DestinationOutputProfileColorComponents);
        AddAttribute(attributes, "destinationOutputProfileAlternateColorSpace", outputIntent.DestinationOutputProfileAlternateColorSpace);
        AddAttribute(attributes, "destinationOutputProfileFilter", outputIntent.DestinationOutputProfileFilter);
        AddAttribute(attributes, "destinationOutputProfileSizeBytes", outputIntent.DestinationOutputProfileSizeBytes);
        AddAttribute(attributes, "destinationOutputProfileDeclaredSizeBytes", outputIntent.DestinationOutputProfileDeclaredSizeBytes);
        AddAttribute(attributes, "destinationOutputProfileColorSpace", outputIntent.DestinationOutputProfileColorSpace);
        AddAttribute(attributes, "destinationOutputProfileHasIccSignature", ToMetadataText(outputIntent.DestinationOutputProfileHasIccSignature));

        return new OfficeDocumentMetadataEntry {
            Id = id,
            Category = "pdf.outputIntent",
            Name = outputIntent.Subtype ?? "OutputIntent",
            Value = outputIntent.OutputConditionIdentifier ?? outputIntent.OutputCondition ?? outputIntent.Info,
            ValueType = "object",
            SourceObjectId = outputIntent.ObjectNumber.HasValue
                ? outputIntent.ObjectNumber.Value.ToString(CultureInfo.InvariantCulture)
                : null,
            Attributes = attributes
        };
    }

    private static void AddOptionalContentMetadata(List<OfficeDocumentMetadataEntry> entries, PdfOptionalContentProperties? optionalContent) {
        if (optionalContent == null || optionalContent.GroupCount == 0) {
            return;
        }

        int initiallyVisibleCount = 0;
        int initiallyHiddenCount = 0;
        int lockedCount = 0;
        int orderedCount = 0;
        for (int i = 0; i < optionalContent.Groups.Count; i++) {
            PdfOptionalContentGroup group = optionalContent.Groups[i];
            if (group.IsInitiallyVisible == true) {
                initiallyVisibleCount++;
            } else if (group.IsInitiallyVisible == false) {
                initiallyHiddenCount++;
            }

            if (group.IsLocked) {
                lockedCount++;
            }

            if (group.IsInDefaultOrder) {
                orderedCount++;
            }
        }

        AddCountMetadata(entries, "pdf-optional-content-group-count", "pdf.optionalContent", "GroupCount", optionalContent.GroupCount);
        AddCountMetadata(entries, "pdf-optional-content-initially-visible-count", "pdf.optionalContent", "InitiallyVisibleCount", initiallyVisibleCount);
        AddCountMetadata(entries, "pdf-optional-content-initially-hidden-count", "pdf.optionalContent", "InitiallyHiddenCount", initiallyHiddenCount);
        AddCountMetadata(entries, "pdf-optional-content-locked-count", "pdf.optionalContent", "LockedCount", lockedCount);
        AddCountMetadata(entries, "pdf-optional-content-default-order-count", "pdf.optionalContent", "DefaultOrderCount", orderedCount);

        entries.Add(BuildOptionalContentConfigurationMetadataEntry(optionalContent));
        for (int i = 0; i < optionalContent.Groups.Count; i++) {
            entries.Add(BuildOptionalContentGroupMetadataEntry(optionalContent.Groups[i], i));
        }
    }

    private static OfficeDocumentMetadataEntry BuildOptionalContentConfigurationMetadataEntry(PdfOptionalContentProperties optionalContent) {
        var attributes = new Dictionary<string, string>(StringComparer.Ordinal);
        AddAttribute(attributes, "name", optionalContent.DefaultConfigurationName);
        AddAttribute(attributes, "creator", optionalContent.DefaultConfigurationCreator);
        AddAttribute(attributes, "baseState", optionalContent.BaseState);
        AddAttribute(attributes, "onGroupObjectNumbers", FormatPdfIntegerComponents(optionalContent.OnGroupObjectNumbers));
        AddAttribute(attributes, "offGroupObjectNumbers", FormatPdfIntegerComponents(optionalContent.OffGroupObjectNumbers));
        AddAttribute(attributes, "lockedGroupObjectNumbers", FormatPdfIntegerComponents(optionalContent.LockedGroupObjectNumbers));
        AddAttribute(attributes, "orderGroupObjectNumbers", FormatPdfIntegerComponents(optionalContent.OrderGroupObjectNumbers));

        return new OfficeDocumentMetadataEntry {
            Id = "pdf-optional-content-configuration",
            Category = "pdf.optionalContent",
            Name = "DefaultConfiguration",
            Value = optionalContent.DefaultConfigurationName,
            ValueType = "object",
            Attributes = attributes
        };
    }

    private static OfficeDocumentMetadataEntry BuildOptionalContentGroupMetadataEntry(PdfOptionalContentGroup group, int groupIndex) {
        string id = "pdf-optional-content-group-" + groupIndex.ToString("D4", CultureInfo.InvariantCulture);
        var attributes = new Dictionary<string, string>(StringComparer.Ordinal) {
            ["name"] = group.Name,
            ["isLocked"] = ToMetadataText(group.IsLocked),
            ["isInDefaultOrder"] = ToMetadataText(group.IsInDefaultOrder)
        };

        AddAttribute(attributes, "objectNumber", group.ObjectNumber);
        AddAttribute(attributes, "intents", FormatPdfStringComponents(group.Intents));
        AddAttribute(attributes, "isInitiallyVisible", ToMetadataText(group.IsInitiallyVisible));
        AddAttribute(attributes, "viewState", group.ViewState);
        AddAttribute(attributes, "printState", group.PrintState);
        AddAttribute(attributes, "exportState", group.ExportState);
        AddAttribute(attributes, "usageCreator", group.UsageCreator);
        AddAttribute(attributes, "usageSubtype", group.UsageSubtype);

        return new OfficeDocumentMetadataEntry {
            Id = id,
            Category = "pdf.optionalContent.group",
            Name = group.Name,
            Value = group.IsInitiallyVisible.HasValue ? ToMetadataText(group.IsInitiallyVisible.Value) : null,
            ValueType = "object",
            SourceObjectId = group.ObjectNumber.HasValue
                ? group.ObjectNumber.Value.ToString(CultureInfo.InvariantCulture)
                : null,
            Attributes = attributes
        };
    }

    private static void AddAnnotationMetadata(List<OfficeDocumentMetadataEntry> entries, SourceMetadata source, IReadOnlyList<PdfLogicalPage> pages) {
        int annotationCount = 0;
        int annotationGeometryCount = 0;
        int annotationAppearanceCount = 0;
        int visualStyleMetadataCount = 0;
        int pathGeometryMetadataCount = 0;
        int freeTextAppearanceMetadataCount = 0;
        var subtypeCounts = new Dictionary<string, int>(StringComparer.Ordinal);
        var annotationEntries = new List<OfficeDocumentMetadataEntry>();

        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            PdfLogicalPage page = pages[pageIndex];
            IReadOnlyList<PdfAnnotation> annotations = page.Annotations;
            for (int annotationIndex = 0; annotationIndex < annotations.Count; annotationIndex++) {
                PdfAnnotation annotation = annotations[annotationIndex];
                annotationCount++;
                if (annotation.Width > 0D && annotation.Height > 0D) {
                    annotationGeometryCount++;
                }

                if (annotation.HasNormalAppearance) {
                    annotationAppearanceCount++;
                }

                if (annotation.HasVisualStyleMetadata) {
                    visualStyleMetadataCount++;
                }

                if (annotation.HasPathGeometryMetadata) {
                    pathGeometryMetadataCount++;
                }

                string subtypeKey = GetActionTypeKey(annotation.Subtype);
                subtypeCounts[subtypeKey] = subtypeCounts.TryGetValue(subtypeKey, out int subtypeCount) ? subtypeCount + 1 : 1;

                if (annotation.HasFreeTextAppearanceMetadata) {
                    freeTextAppearanceMetadataCount++;
                }

                annotationEntries.Add(BuildAnnotationMetadataEntry(source, page, pageIndex, annotation, annotationIndex));
            }
        }

        AddCountMetadata(entries, "pdf-annotation-count", "pdf.annotation", "Count", annotationCount);
        AddCountMetadata(entries, "pdf-annotation-geometry-count", "pdf.annotation", "GeometryCount", annotationGeometryCount);
        AddCountMetadata(entries, "pdf-annotation-normal-appearance-count", "pdf.annotation", "NormalAppearanceCount", annotationAppearanceCount);
        AddCountMetadata(entries, "pdf-annotation-visual-style-metadata-count", "pdf.annotation", "VisualStyleMetadataCount", visualStyleMetadataCount);
        AddCountMetadata(entries, "pdf-annotation-path-geometry-metadata-count", "pdf.annotation", "PathGeometryMetadataCount", pathGeometryMetadataCount);
        AddCountMetadata(entries, "pdf-annotation-freetext-appearance-metadata-count", "pdf.annotation.freeText", "AppearanceMetadataCount", freeTextAppearanceMetadataCount);
        if (annotationCount > 0) {
            AddNumberMetadata(entries, "pdf-annotation-geometry-coverage", "pdf.annotation", "GeometryCoverage", (double)annotationGeometryCount / annotationCount);
            AddNumberMetadata(entries, "pdf-annotation-normal-appearance-coverage", "pdf.annotation", "NormalAppearanceCoverage", (double)annotationAppearanceCount / annotationCount);
        }

        foreach (KeyValuePair<string, int> subtypeCount in subtypeCounts.OrderBy(item => item.Key, StringComparer.Ordinal)) {
            AddCountMetadata(
                entries,
                "pdf-annotation-" + subtypeCount.Key + "-count",
                "pdf.annotation",
                ToMetadataDisplayName(subtypeCount.Key) + "Count",
                subtypeCount.Value);
        }

        entries.AddRange(annotationEntries);
    }

    private static OfficeDocumentMetadataEntry BuildAnnotationMetadataEntry(SourceMetadata source, PdfLogicalPage page, int pageIndex, PdfAnnotation annotation, int annotationIndex) {
        string subtypeKey = GetActionTypeKey(annotation.Subtype);
        string id = "pdf-page-" + page.PageNumber.ToString("D4", CultureInfo.InvariantCulture) +
            "-selection-" + pageIndex.ToString("D4", CultureInfo.InvariantCulture) +
            "-annotation-" + annotationIndex.ToString("D4", CultureInfo.InvariantCulture);
        var attributes = new Dictionary<string, string>(StringComparer.Ordinal) {
            ["subtype"] = annotation.Subtype,
            ["hasNormalAppearance"] = ToMetadataText(annotation.HasNormalAppearance),
            ["x"] = annotation.X1.ToString("R", CultureInfo.InvariantCulture),
            ["y"] = annotation.Y1.ToString("R", CultureInfo.InvariantCulture),
            ["width"] = annotation.Width.ToString("R", CultureInfo.InvariantCulture),
            ["height"] = annotation.Height.ToString("R", CultureInfo.InvariantCulture)
        };

        AddAttribute(attributes, "objectNumber", annotation.ObjectNumber);
        AddAttribute(attributes, "flags", annotation.Flags);
        AddAttribute(attributes, "name", annotation.Name);
        AddAttribute(attributes, "title", annotation.Title);
        AddAttribute(attributes, "modified", annotation.Modified);
        AddAttribute(attributes, "color", FormatPdfColorComponents(annotation.Color));
        AddAttribute(attributes, "interiorColor", FormatPdfColorComponents(annotation.InteriorColor));
        AddAttribute(attributes, "opacity", annotation.Opacity);
        AddAttribute(attributes, "borderWidth", annotation.BorderWidth);
        AddAttribute(attributes, "borderStyle", annotation.BorderStyle);
        AddAttribute(attributes, "borderDashPattern", FormatPdfColorComponents(annotation.BorderDashPattern));
        AddAttribute(attributes, "borderEffectStyle", annotation.BorderEffectStyle);
        AddAttribute(attributes, "borderEffectIntensity", annotation.BorderEffectIntensity);
        AddAttribute(attributes, "rectangleDifferences", FormatPdfColorComponents(annotation.RectangleDifferences));
        AddAttribute(attributes, "calloutLine", FormatPdfColorComponents(annotation.CalloutLine));
        AddAttribute(attributes, "calloutLineEnding", annotation.CalloutLineEnding);
        AddAttribute(attributes, "lineStartEnding", annotation.LineStartEnding);
        AddAttribute(attributes, "lineEndEnding", annotation.LineEndEnding);
        AddAttribute(attributes, "quadPoints", FormatPdfColorComponents(annotation.QuadPoints));
        AddAttribute(attributes, "lineCoordinates", FormatPdfColorComponents(annotation.LineCoordinates));
        AddAttribute(attributes, "vertices", FormatPdfColorComponents(annotation.Vertices));
        AddAttribute(attributes, "inkList", FormatNestedPdfNumberLists(annotation.InkList));
        AddAttribute(attributes, "defaultAppearance", annotation.DefaultAppearance);
        AddAttribute(attributes, "defaultStyle", annotation.DefaultStyle);
        AddAttribute(attributes, "richContentsPlainText", annotation.RichContentsPlainText);
        AddAttribute(attributes, "effectiveFontSize", annotation.EffectiveFontSize);
        AddAttribute(attributes, "effectiveTextColor", annotation.EffectiveTextColor.HasValue ? FormatPdfColor(annotation.EffectiveTextColor.Value) : null);
        AddAttribute(attributes, "effectiveTextAlign", annotation.EffectiveTextAlign?.ToString());

        return new OfficeDocumentMetadataEntry {
            Id = id,
            Category = string.Equals(annotation.Subtype, "FreeText", StringComparison.Ordinal) && annotation.HasFreeTextAppearanceMetadata
                ? "pdf.annotation.freeText"
                : "pdf.annotation",
            Name = annotation.Subtype,
            Value = annotation.Contents ?? annotation.RichContentsPlainText,
            ValueType = "object",
            SourceObjectId = annotation.ObjectNumber?.ToString(CultureInfo.InvariantCulture),
            Location = BuildLocation(source, page.PageNumber, pageIndex, "annotation", id),
            Attributes = attributes
        };
    }

    private static void AddTableQualityMetadata(List<OfficeDocumentMetadataEntry> entries, IReadOnlyList<ReaderTable> tables) {
        int geometryCount = 0;
        int lowConfidenceCount = 0;
        int numericColumnCount = 0;
        int fallbackColumnNameCount = 0;
        int missingCellCount = 0;
        for (int tableIndex = 0; tableIndex < tables.Count; tableIndex++) {
            ReaderTable table = tables[tableIndex];
            ReaderTableDiagnostics? diagnostics = table.Diagnostics;
            if (diagnostics != null) {
                if (diagnostics.HasGeometry) {
                    geometryCount++;
                }

                if (diagnostics.Confidence < HighConfidenceTableThreshold) {
                    lowConfidenceCount++;
                }

                missingCellCount += diagnostics.MissingCellCount;
            }

            for (int profileIndex = 0; profileIndex < table.ColumnProfiles.Count; profileIndex++) {
                ReaderTableColumnProfile profile = table.ColumnProfiles[profileIndex];
                if (profile.IsNumeric) {
                    numericColumnCount++;
                }

                if (IsFallbackColumnName(profile.Name, profile.Index)) {
                    fallbackColumnNameCount++;
                }
            }
        }

        AddCountMetadata(entries, "pdf-table-count", "pdf.table", "Count", tables.Count);
        AddCountMetadata(entries, "pdf-table-geometry-count", "pdf.table", "GeometryCount", geometryCount);
        AddCountMetadata(entries, "pdf-table-low-confidence-count", "pdf.table", "LowConfidenceCount", lowConfidenceCount);
        AddCountMetadata(entries, "pdf-table-numeric-column-count", "pdf.table", "NumericColumnCount", numericColumnCount);
        AddCountMetadata(entries, "pdf-table-fallback-column-name-count", "pdf.table", "FallbackColumnNameCount", fallbackColumnNameCount);
        AddCountMetadata(entries, "pdf-table-missing-cell-count", "pdf.table", "MissingCellCount", missingCellCount);
        if (tables.Count > 0) {
            AddNumberMetadata(entries, "pdf-table-geometry-coverage", "pdf.table", "GeometryCoverage", (double)geometryCount / tables.Count);
        }
    }

    private static void AddOcrCandidateMetadata(List<OfficeDocumentMetadataEntry> entries, IReadOnlyList<OfficeDocumentOcrCandidate> ocrCandidates) {
        int imageCandidateCount = 0;
        int pageCandidateCount = 0;
        int assetLinkedCount = 0;
        int geometryCount = 0;
        int candidateImageCount = 0;
        int candidateTextBlockCount = 0;
        for (int i = 0; i < ocrCandidates.Count; i++) {
            OfficeDocumentOcrCandidate candidate = ocrCandidates[i];
            if (string.Equals(candidate.Kind, "image", StringComparison.Ordinal)) {
                imageCandidateCount++;
            } else if (string.Equals(candidate.Kind, "page", StringComparison.Ordinal)) {
                pageCandidateCount++;
            }

            if (!string.IsNullOrWhiteSpace(candidate.AssetId)) {
                assetLinkedCount++;
            }

            if (candidate.Region != null && candidate.Region.Width > 0D && candidate.Region.Height > 0D) {
                geometryCount++;
            }

            candidateImageCount += candidate.ImageCount ?? 0;
            candidateTextBlockCount += candidate.TextBlockCount ?? 0;
        }

        AddCountMetadata(entries, "pdf-ocr-candidate-count", "pdf.ocr", "CandidateCount", ocrCandidates.Count);
        AddCountMetadata(entries, "pdf-ocr-image-candidate-count", "pdf.ocr", "ImageCandidateCount", imageCandidateCount);
        AddCountMetadata(entries, "pdf-ocr-page-candidate-count", "pdf.ocr", "PageCandidateCount", pageCandidateCount);
        AddCountMetadata(entries, "pdf-ocr-asset-linked-count", "pdf.ocr", "AssetLinkedCount", assetLinkedCount);
        AddCountMetadata(entries, "pdf-ocr-candidate-geometry-count", "pdf.ocr", "CandidateGeometryCount", geometryCount);
        AddCountMetadata(entries, "pdf-ocr-candidate-image-count", "pdf.ocr", "CandidateImageCount", candidateImageCount);
        if (candidateTextBlockCount > 0) {
            AddCountMetadata(entries, "pdf-ocr-candidate-text-block-count", "pdf.ocr", "CandidateTextBlockCount", candidateTextBlockCount);
        }

        if (ocrCandidates.Count > 0) {
            AddNumberMetadata(entries, "pdf-ocr-candidate-geometry-coverage", "pdf.ocr", "CandidateGeometryCoverage", (double)geometryCount / ocrCandidates.Count);
        }
    }

    private static void AddActionMetadata(List<OfficeDocumentMetadataEntry> entries, PdfLogicalDocument document, IReadOnlyList<PdfLogicalPage> selectedPages) {
        IReadOnlyList<ReaderActionSummary>? actions = BuildActions(document, selectedPages, page: null, includeDocumentActions: true);
        if (actions == null || actions.Count == 0) {
            return;
        }

        int activeActionCount = 0;
        int chainedActionCount = 0;
        int potentiallyUnsafeActionCount = 0;
        var scopeCounts = new Dictionary<string, int>(StringComparer.Ordinal);
        var typeCounts = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int i = 0; i < actions.Count; i++) {
            ReaderActionSummary action = actions[i];
            string scope = GetActionScopeKey(action.Scope);
            scopeCounts[scope] = scopeCounts.TryGetValue(scope, out int scopeCount) ? scopeCount + 1 : 1;

            if (action.Scope != ReaderActionScope.DocumentOpen) {
                activeActionCount++;
            }

            if (action.IsChainedAction) {
                chainedActionCount++;
            }

            if (action.IsPotentiallyUnsafe) {
                potentiallyUnsafeActionCount++;
            }

            string type = GetActionTypeKey(action.ActionType);
            typeCounts[type] = typeCounts.TryGetValue(type, out int typeCount) ? typeCount + 1 : 1;
        }

        AddCountMetadata(entries, "pdf-action-count", "pdf.action", "Count", actions.Count);
        AddCountMetadata(entries, "pdf-active-action-count", "pdf.action", "ActiveCount", activeActionCount);
        AddCountMetadata(entries, "pdf-action-chained-count", "pdf.action", "ChainedCount", chainedActionCount);
        AddCountMetadata(entries, "pdf-action-potentially-unsafe-count", "pdf.action", "PotentiallyUnsafeCount", potentiallyUnsafeActionCount);

        foreach (KeyValuePair<string, int> scopeCount in scopeCounts.OrderBy(item => item.Key, StringComparer.Ordinal)) {
            AddCountMetadata(
                entries,
                "pdf-action-" + scopeCount.Key + "-count",
                "pdf.action",
                ToMetadataDisplayName(scopeCount.Key) + "Count",
                scopeCount.Value);
        }

        foreach (KeyValuePair<string, int> typeCount in typeCounts.OrderBy(item => item.Key, StringComparer.Ordinal)) {
            AddCountMetadata(
                entries,
                "pdf-action-type-" + typeCount.Key + "-count",
                "pdf.action",
                ToMetadataDisplayName(typeCount.Key) + "ActionCount",
                typeCount.Value);
        }
    }

    private static string GetActionScopeKey(ReaderActionScope scope) {
        switch (scope) {
            case ReaderActionScope.DocumentOpen:
                return "document-open";
            case ReaderActionScope.Catalog:
                return "catalog";
            case ReaderActionScope.Page:
                return "page";
            case ReaderActionScope.Annotation:
                return "annotation";
            default:
                return "unknown";
        }
    }

    private static string GetActionTypeKey(string actionType) {
        if (string.IsNullOrWhiteSpace(actionType)) {
            return "unknown";
        }

        var builder = new System.Text.StringBuilder(actionType.Length);
        for (int i = 0; i < actionType.Length; i++) {
            char character = actionType[i];
            if (char.IsLetterOrDigit(character)) {
                builder.Append(char.ToLowerInvariant(character));
            } else if (builder.Length > 0 && builder[builder.Length - 1] != '-') {
                builder.Append('-');
            }
        }

        if (builder.Length > 0 && builder[builder.Length - 1] == '-') {
            builder.Length--;
        }

        return builder.Length == 0 ? "unknown" : builder.ToString();
    }

    private static string? FormatPdfColorComponents(IReadOnlyList<double> components) {
        if (components.Count == 0) {
            return null;
        }

        var builder = new System.Text.StringBuilder();
        for (int i = 0; i < components.Count; i++) {
            if (i > 0) {
                builder.Append(',');
            }

            builder.Append(components[i].ToString("R", CultureInfo.InvariantCulture));
        }

        return builder.ToString();
    }

    private static string FormatPdfColor(PdfColor color) =>
        color.R.ToString("R", CultureInfo.InvariantCulture) + "," +
        color.G.ToString("R", CultureInfo.InvariantCulture) + "," +
        color.B.ToString("R", CultureInfo.InvariantCulture);

    private static string? FormatNestedPdfNumberLists(IReadOnlyList<IReadOnlyList<double>> paths) {
        if (paths.Count == 0) {
            return null;
        }

        var builder = new System.Text.StringBuilder();
        for (int i = 0; i < paths.Count; i++) {
            if (i > 0) {
                builder.Append(';');
            }

            builder.Append(FormatPdfColorComponents(paths[i]));
        }

        return builder.ToString();
    }

    private static string? FormatPdfIntegerComponents(IReadOnlyList<int> components) {
        if (components.Count == 0) {
            return null;
        }

        var builder = new System.Text.StringBuilder();
        for (int i = 0; i < components.Count; i++) {
            if (i > 0) {
                builder.Append(',');
            }

            builder.Append(components[i].ToString(CultureInfo.InvariantCulture));
        }

        return builder.ToString();
    }

    private static string? FormatPdfStringComponents(IReadOnlyList<string> components) {
        if (components.Count == 0) {
            return null;
        }

        var builder = new System.Text.StringBuilder();
        for (int i = 0; i < components.Count; i++) {
            if (string.IsNullOrWhiteSpace(components[i])) {
                continue;
            }

            if (builder.Length > 0) {
                builder.Append(',');
            }

            builder.Append(components[i]);
        }

        return builder.Length == 0 ? null : builder.ToString();
    }

    private static void AddFormWidgetMetadata(List<OfficeDocumentMetadataEntry> entries, IReadOnlyList<PdfLogicalPage> pages) {
        int widgetCount = 0;
        int widgetGeometryCount = 0;
        var selectedFields = new HashSet<PdfFormField>();
        var kindCounts = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            IReadOnlyList<PdfLogicalFormWidget> widgets = pages[pageIndex].FormWidgets;
            for (int widgetIndex = 0; widgetIndex < widgets.Count; widgetIndex++) {
                PdfLogicalFormWidget widget = widgets[widgetIndex];
                widgetCount++;
                if (widget.Width > 0D && widget.Height > 0D) {
                    widgetGeometryCount++;
                }

                if (selectedFields.Add(widget.Field)) {
                    string kind = GetFormFieldKind(widget.Field.Kind);
                    kindCounts[kind] = kindCounts.TryGetValue(kind, out int count) ? count + 1 : 1;
                }
            }
        }

        AddCountMetadata(entries, "pdf-form-widget-count", "pdf.form", "WidgetCount", widgetCount);
        AddCountMetadata(entries, "pdf-form-widget-geometry-count", "pdf.form", "WidgetGeometryCount", widgetGeometryCount);
        if (widgetCount > 0) {
            AddNumberMetadata(entries, "pdf-form-widget-geometry-coverage", "pdf.form", "WidgetGeometryCoverage", (double)widgetGeometryCount / widgetCount);
        }

        foreach (KeyValuePair<string, int> kindCount in kindCounts.OrderBy(item => item.Key, StringComparer.Ordinal)) {
            AddCountMetadata(
                entries,
                "pdf-form-" + kindCount.Key + "-count",
                "pdf.form",
                ToMetadataDisplayName(kindCount.Key) + "Count",
                kindCount.Value);
        }
    }

    private static string GetFormFieldKind(PdfFormFieldKind kind) {
        switch (kind) {
            case PdfFormFieldKind.Text:
                return "text";
            case PdfFormFieldKind.Button:
                return "button";
            case PdfFormFieldKind.Choice:
                return "choice";
            case PdfFormFieldKind.Signature:
                return "signature";
            default:
                return "unknown";
        }
    }

    private static string ToMetadataDisplayName(string value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return string.Empty;
        }

        var builder = new System.Text.StringBuilder(value.Length);
        bool upperNext = true;
        for (int i = 0; i < value.Length; i++) {
            char character = value[i];
            if (character == '-' || character == '_' || character == ' ') {
                upperNext = true;
                continue;
            }

            builder.Append(upperNext ? char.ToUpperInvariant(character) : character);
            upperNext = false;
        }

        return builder.ToString();
    }
}
