using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.PowerPoint.LegacyPpt.Write;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        private const double EmusPerLegacyMasterUnit = 1587.5d;

        /// <summary>Loads a binary `.ppt`, `.pot`, or `.pps` file into the normal editable PowerPoint model.</summary>
        internal static PowerPointDocument LoadLegacyPpt(string path, LegacyPptImportOptions? options = null) {
            ArgumentNullException.ThrowIfNull(path);
            LegacyPptDocument legacy = LegacyPptDocument.Load(path, options);
            return ProjectLoadedLegacyPpt(legacy, path,
                PowerPointDocumentLoadRouting.GetFormat(path, legacyDefault: true), new PowerPointLoadOptions());
        }

        /// <summary>Loads a binary PowerPoint stream into the normal editable PowerPoint model.</summary>
        internal static PowerPointDocument LoadLegacyPpt(Stream stream, LegacyPptImportOptions? options = null) {
            LegacyPptDocument legacy = LegacyPptDocument.Load(stream, options);
            return ProjectLoadedLegacyPpt(legacy, sourcePath: null, PowerPointFileFormat.Ppt,
                new PowerPointLoadOptions());
        }

        /// <summary>Loads a binary PowerPoint file and returns its projected presentation and import report.</summary>
        internal static LegacyPptLoadResult LoadLegacyPptWithReport(string path, LegacyPptImportOptions? options = null) {
            ArgumentNullException.ThrowIfNull(path);
            LegacyPptDocument legacy = LegacyPptDocument.Load(path, options);
            try {
                return new LegacyPptLoadResult(ProjectLoadedLegacyPpt(legacy, path,
                    PowerPointDocumentLoadRouting.GetFormat(path, legacyDefault: true), new PowerPointLoadOptions()), legacy);
            } catch (InvalidDataException exception) {
                return new LegacyPptLoadResult(document: null, legacy, exception);
            }
        }

        /// <summary>Loads a binary PowerPoint stream and returns its projected presentation and import report.</summary>
        internal static LegacyPptLoadResult LoadLegacyPptWithReport(Stream stream, LegacyPptImportOptions? options = null) {
            LegacyPptDocument legacy = LegacyPptDocument.Load(stream, options);
            try {
                return new LegacyPptLoadResult(ProjectLoadedLegacyPpt(legacy, sourcePath: null,
                    PowerPointFileFormat.Ppt, new PowerPointLoadOptions()), legacy);
            } catch (InvalidDataException exception) {
                return new LegacyPptLoadResult(document: null, legacy, exception);
            }
        }

        private static PowerPointDocument LoadLegacyPptFromNormalFlow(byte[] bytes, string? sourcePath,
            Stream? sourceStream, PowerPointLoadOptions options,
            CancellationToken cancellationToken = default) {
            if (options.PersistenceMode == DocumentPersistenceMode.SaveOnDispose && sourceStream == null
                && string.IsNullOrEmpty(sourcePath)) {
                throw new NotSupportedException("SaveOnDispose requires an associated destination for binary PowerPoint sources.");
            }
            LegacyPptDocument legacy = LegacyPptDocument.Load(bytes,
                options.LegacyPptImportOptions, cancellationToken);
            PowerPointFileFormat sourceFormat = PowerPointDocumentLoadRouting.GetFormat(sourcePath, legacyDefault: true);
            return ProjectLoadedLegacyPpt(legacy, sourcePath, sourceFormat,
                options, sourceStream, cancellationToken);
        }

        private static PowerPointDocument LoadEncryptedLegacyPptFromNormalFlow(
            byte[] bytes, string password, PowerPointFileFormat sourceFormat,
            PowerPointLoadOptions options,
            CancellationToken cancellationToken = default) {
            LegacyPptImportOptions sourceOptions = options.LegacyPptImportOptions
                ?? new LegacyPptImportOptions();
            if (!options.AllowLegacyProjection) throw new NotSupportedException("This load profile accepts Open XML presentations only; encrypted binary PPT projection is disabled.");
            var importOptions = new LegacyPptImportOptions {
                MaxInputBytes = sourceOptions.MaxInputBytes,
                ReportUnsupportedContent = sourceOptions.ReportUnsupportedContent,
                MaxRecordCount = sourceOptions.MaxRecordCount,
                MaxRecordDepth = sourceOptions.MaxRecordDepth,
                MaxMasterCount = sourceOptions.MaxMasterCount,
                MaxConnectorRuleCount = sourceOptions.MaxConnectorRuleCount,
                MaxCommentCount = sourceOptions.MaxCommentCount,
                MaxTextStyle9EntryCount =
                    sourceOptions.MaxTextStyle9EntryCount,
                MaxDecodedStorageBytes =
                    sourceOptions.MaxDecodedStorageBytes,
                Password = password
            };
            LegacyPptDocument legacy = LegacyPptDocument.Load(bytes,
                importOptions, cancellationToken);
            return WithSourcePackageInspection(ProjectLoadedLegacyPpt(legacy, sourcePath: null,
                sourceFormat, options,
                cancellationToken: cancellationToken), bytes, options, "encrypted-source-container");
        }


        internal static PowerPointDocument ProjectLoadedLegacyPpt(LegacyPptDocument legacy,
            string? sourcePath, PowerPointFileFormat sourceFormat, PowerPointLoadOptions loadOptions,
            Stream? sourceStream = null,
            CancellationToken cancellationToken = default) {
            ArgumentNullException.ThrowIfNull(legacy);
            cancellationToken.ThrowIfCancellationRequested();
            ValidateLegacyContentSecurity(legacy, loadOptions);
            using PowerPointDocument projected = CreateForImport();
            LegacyPptImportParts parts = LegacyPptImportParts.For(projected);
            if (!parts.IsPackageHeld) {
                throw new InvalidOperationException("The scratch presentation of a binary import must start on its own package.");
            }

            // Everything that adds parts runs while the scratch presentation is on its own package: document settings,
            // masters, layouts, themes, special masters, the VBA project, slides with their designs, custom shows,
            // master interactions that wait for the slides, and comments.
            ApplyLegacyDocumentSettings(projected, parts, legacy);
            var soundContext = new LegacyPptSoundProjectionContext(legacy);
            var deferredInteractions =
                new List<LegacyPptDeferredProjection>();
            LegacyPptLayoutCatalog layoutTargets = ProjectLegacyMasters(parts,
                legacy, soundContext, deferredInteractions);
            ProjectLegacySpecialMasters(parts, legacy, soundContext,
                deferredInteractions);
            ProjectLegacyVbaProject(parts, legacy);

            var slideProjections = new List<(LegacyPptSlide Source, PowerPointSlide Target)>(
                legacy.Slides.Count);
            var slidePartsByLegacyId = new LegacyPptSlidePartMap();
            foreach (LegacyPptSlide legacySlide in legacy.Slides) {
                cancellationToken.ThrowIfCancellationRequested();
                PowerPointSlide slide = layoutTargets.TryGet(legacySlide,
                    out LegacyPptLayoutTarget target)
                    ? projected.AddSlide(target.MasterIndex, target.LayoutIndex)
                    : projected.AddSlide();
                slide.Hidden = legacySlide.Hidden;
                ProjectLegacySlideDesign(parts, slide, legacySlide);
                slideProjections.Add((legacySlide, slide));
                if (slidePartsByLegacyId.ContainsKey(legacySlide.SlideId)) {
                    throw new InvalidDataException(
                        $"Binary slide identifier {legacySlide.SlideId} occurs more than once.");
                }
                slidePartsByLegacyId.Add(legacySlide.SlideId, slide);
            }

            ProjectLegacyCustomShows(parts, legacy, slidePartsByLegacyId);
            foreach (LegacyPptDeferredProjection interaction in
                     deferredInteractions) {
                interaction.Apply(slidePartsByLegacyId);
            }
            ProjectLegacyComments(parts, legacy,
                slideProjections.Select(item => item.Target.PartRef).ToList());

            // From here the engine may open the SDK (titles, OLE objects, media, tables, notes); parts the slides still
            // need are written into the saved package.
            ProjectLegacyLateTables(parts, slidePartsByLegacyId, soundContext);
            foreach ((LegacyPptSlide legacySlide, PowerPointSlide slide) in slideProjections) {
                cancellationToken.ThrowIfCancellationRequested();
                ProjectLegacyTransition(slide, legacySlide.Transition, soundContext);
                if (HasLegacyGroupedTable(legacySlide.Shapes)) {
                    // Grouped tables are built through the SDK, so it opens before the group's tree is read.
                    EnsureLegacySdk(projected);
                }
                var projectedShapeIds = new Dictionary<uint, uint>();
                foreach (LegacyPptShape shape in legacySlide.Shapes) {
                    cancellationToken.ThrowIfCancellationRequested();
                    ProjectLegacyShape(slide, shape,
                        slidePartsByLegacyId, soundContext,
                        projectedShapeIds);
                }
                ProjectLegacyConnectorRules(slide, legacySlide.ConnectorRules, projectedShapeIds);
                ProjectLegacyAnimations(slide, legacySlide.Shapes,
                    projectedShapeIds, soundContext);
                if (legacySlide.NotesPage != null) {
                    ProjectLegacyNotesPageThroughSdk(slide, legacySlide.NotesPage,
                        slidePartsByLegacyId, soundContext);
                } else if (!string.IsNullOrWhiteSpace(legacySlide.NotesText)) {
                    slide.Notes.Text = legacySlide.NotesText;
                }
            }

            cancellationToken.ThrowIfCancellationRequested();
            bool readOnlyProjection = loadOptions.AccessMode ==
                DocumentAccessMode.ReadOnly;
            if (readOnlyProjection) {
                LegacyPptPropertySetCodec.Apply(projected, legacy.Package);
            }

            byte[] packageBytes = CreateLegacyProjectedPackageBytes(projected);
            packageBytes = parts.ApplyPending(packageBytes);
            cancellationToken.ThrowIfCancellationRequested();
            if (legacy.VbaProject != null) {
                packageBytes = ConvertProjectedVbaPackageToMacroEnabled(
                    packageBytes, loadOptions);
            }
            PowerPointDocument presentation = LoadPackage(packageBytes,
                sourcePath, sourceStream, loadOptions, captureSourcePackageSecurity: false);
            try {
                cancellationToken.ThrowIfCancellationRequested();
            } catch {
                presentation.Dispose();
                throw;
            }
            if (readOnlyProjection) {
                LegacyPptPropertySetCodec.ApplyReadOnlyDateOverrides(
                    presentation, legacy.Package);
            } else {
                LegacyPptPropertySetCodec.Apply(presentation, legacy.Package);
            }
            LegacyPptProjectionMap projectionMap = LegacyPptProjectionMap.Create(presentation, legacy);
            presentation.MarkLoadedFromLegacyPpt(sourcePath, legacy, projectionMap, sourceFormat);
            return presentation;
        }

        /// <summary>
        /// The projected slides by binary slide id. Each lookup reads the slide's part as the scratch presentation holds it
        /// then, a part of its package or, once the SDK opened, the SDK part.
        /// </summary>
        private sealed class LegacyPptSlidePartMap : IReadOnlyDictionary<uint, PowerPointPartRef> {
            private readonly Dictionary<uint, PowerPointSlide> _slides = new();

            internal void Add(uint slideId, PowerPointSlide slide) => _slides.Add(slideId, slide);

            public PowerPointPartRef this[uint key] => _slides[key].PartRef;
            public IEnumerable<uint> Keys => _slides.Keys;
            public IEnumerable<PowerPointPartRef> Values => _slides.Values.Select(slide => slide.PartRef);
            public int Count => _slides.Count;
            public bool ContainsKey(uint key) => _slides.ContainsKey(key);

            public bool TryGetValue(uint key, out PowerPointPartRef value) {
                if (_slides.TryGetValue(key, out PowerPointSlide? slide)) {
                    value = slide.PartRef;
                    return true;
                }
                value = default;
                return false;
            }

            public IEnumerator<KeyValuePair<uint, PowerPointPartRef>> GetEnumerator() =>
                _slides.Select(item => new KeyValuePair<uint, PowerPointPartRef>(item.Key, item.Value.PartRef)).GetEnumerator();

            System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
        }

        private static bool HasLegacyGroupedTable(IEnumerable<LegacyPptShape> shapes) =>
            shapes.Any(shape => shape.Kind == LegacyPptShapeKind.Group
                && EnumerateLegacyShapes(shape.Children).Any(child => child.Kind == LegacyPptShapeKind.Table));

        private static OfficeXmlNode ProjectLegacyShape(PowerPointSlide slide, LegacyPptShape shape,
            IReadOnlyDictionary<uint, PowerPointPartRef> slidePartsByLegacyId,
            LegacyPptSoundProjectionContext soundContext,
            IDictionary<uint, uint> projectedShapeIds) {
            long left = ToEmus(shape.Bounds.Left);
            long top = ToEmus(shape.Bounds.Top);
            long width = Math.Max(1L, ToEmus(shape.Bounds.Width));
            long height = Math.Max(1L, ToEmus(shape.Bounds.Height));
            PowerPointShape? projectedShape = null;
            OfficeXmlNode projectedElement = default;
            switch (shape.Kind) {
                case LegacyPptShapeKind.TextBox:
                    PowerPointTextBox textBox = shape.PlaceholderKind == LegacyPptPlaceholderKind.Title
                        || shape.PlaceholderKind == LegacyPptPlaceholderKind.CenterTitle
                        || shape.PlaceholderKind == LegacyPptPlaceholderKind.VerticalTitle
                        ? slide.AddTitle(shape.Text,global::OfficeIMO.OfficeLength.Emus(left),global::OfficeIMO.OfficeLength.Emus(top),global::OfficeIMO.OfficeLength.Emus(width),global::OfficeIMO.OfficeLength.Emus(height))
                        : slide.AddTextBox(shape.Text,global::OfficeIMO.OfficeLength.Emus(left),global::OfficeIMO.OfficeLength.Emus(top),global::OfficeIMO.OfficeLength.Emus(width),global::OfficeIMO.OfficeLength.Emus(height));
                    PowerPointPlaceholderType? placeholder = MapPlaceholder(shape.PlaceholderKind);
                    if (placeholder.HasValue) textBox.PlaceholderType = placeholder.Value;
                    PowerPointPartRef slidePart = slide.PartRef;
                    LegacyPptTextProjection.Apply(textBox.XmlElement,
                        shape.TextBody, shape.TextFrame,
                        CreateLegacyTextInteractionProjector(slidePart, slidePartsByLegacyId, soundContext),
                        pictureBullet => ProjectLegacyPictureBullet(
                            slidePart, pictureBullet));
                    if (shape.OfficeArtShapeType != 202
                        && LegacyPptShapeGeometryMapper.TryGetPreset(shape.OfficeArtShapeType,
                            out OfficePresetShapeType textGeometry)
                        && textBox.XmlElement is { IsNone: false } textShape
                        && textShape.Is(LegacyPptXml.P, "sp")
                        && textShape.Child(LegacyPptXml.P, "spPr").Child(LegacyPptXml.A, "prstGeom") is { IsNone: false } preset) {
                        preset.Set("prst", textGeometry.ToOpenXmlToken());
                    }
                    projectedShape = textBox;
                    break;
                case LegacyPptShapeKind.Rectangle:
                    projectedShape = slide.AddShape(OfficePresetShapeType.Rectangle, OfficeIMO.OfficeLength.Emus(left), OfficeIMO.OfficeLength.Emus(top), OfficeIMO.OfficeLength.Emus(width), OfficeIMO.OfficeLength.Emus(height));
                    break;
                case LegacyPptShapeKind.Ellipse:
                    projectedShape = slide.AddShape(OfficePresetShapeType.Ellipse, OfficeIMO.OfficeLength.Emus(left), OfficeIMO.OfficeLength.Emus(top), OfficeIMO.OfficeLength.Emus(width), OfficeIMO.OfficeLength.Emus(height));
                    break;
                case LegacyPptShapeKind.Line:
                    projectedShape = slide.AddShape(OfficePresetShapeType.Line, OfficeIMO.OfficeLength.Emus(left), OfficeIMO.OfficeLength.Emus(top), OfficeIMO.OfficeLength.Emus(width), OfficeIMO.OfficeLength.Emus(height));
                    break;
                case LegacyPptShapeKind.AutoShape:
                    if (LegacyPptShapeGeometryMapper.TryGetPreset(shape.OfficeArtShapeType,
                            out OfficePresetShapeType geometry)) {
                        projectedShape = slide.AddShape(geometry, OfficeIMO.OfficeLength.Emus(left), OfficeIMO.OfficeLength.Emus(top), OfficeIMO.OfficeLength.Emus(width), OfficeIMO.OfficeLength.Emus(height));
                    }
                    break;
                case LegacyPptShapeKind.Connector:
                    if (LegacyPptShapeGeometryMapper.TryGetPreset(shape.OfficeArtShapeType,
                            out OfficePresetShapeType connectorGeometry)) {
                        projectedElement = slide.AddLegacyConnectionShape(connectorGeometry, left, top, width, height);
                    }
                    break;
                case LegacyPptShapeKind.Picture:
                    if (shape.Picture?.HasImportableImage == true && shape.Picture.ContentType != null) {
                        using var image = new MemoryStream(shape.Picture.ImageBytes, writable: false);
                        projectedShape = slide.AddImage(image,
                            GetLegacyPicturePartType(shape.Picture.ContentType), OfficeLength.Emus(left),
                            OfficeLength.Emus(top), OfficeLength.Emus(width), OfficeLength.Emus(height));
                        projectedShape.Bounds = new PowerPointLayoutBox(OfficeLength.Emus(left), OfficeLength.Emus(top), OfficeLength.Emus(width), OfficeLength.Emus(height));
                    }
                    break;
                case LegacyPptShapeKind.OleObject:
                    if (shape.OleObject != null) {
                        using var storage = new MemoryStream(
                            shape.OleObject.GetBytes(), writable: false);
                        PowerPointOleObject ole = slide.AddOleObject(storage,
                            shape.OleObject.ProgId ?? "Package", OfficeIMO.OfficeLength.Emus(left), OfficeIMO.OfficeLength.Emus(top),
                            OfficeIMO.OfficeLength.Emus(width), OfficeIMO.OfficeLength.Emus(height));
                        ole.ShowAsIcon = shape.OleObject.DrawAspect
                            == LegacyPptOleDrawAspect.Icon;
                        ole.FollowColorScheme = shape.OleObject.ColorFollow switch {
                            LegacyPptOleColorFollow.Scheme =>
                                PowerPointOleFollowColorScheme.Full,
                            LegacyPptOleColorFollow.TextAndBackground =>
                                PowerPointOleFollowColorScheme.TextAndBackground,
                            _ => PowerPointOleFollowColorScheme.None
                        };
                        ApplyLegacyOlePreview(slide.PartRef, ole.XmlElement, shape);
                        projectedShape = ole;
                    }
                    break;
                case LegacyPptShapeKind.Media:
                    if (shape.Media?.HasProjectableAudio == true) {
                        using var media = new MemoryStream(
                            shape.Media.GetData(), writable: false);
                        projectedShape = slide.AddAudio(media, "audio/wav",
                            ".wav", OfficeIMO.OfficeLength.Emus(left), OfficeIMO.OfficeLength.Emus(top), OfficeIMO.OfficeLength.Emus(width), OfficeIMO.OfficeLength.Emus(height));
                    }
                    break;
                case LegacyPptShapeKind.Table:
                    projectedShape = ProjectLegacyTableThroughSdk(slide, shape,
                        slidePartsByLegacyId, soundContext);
                    break;
                case LegacyPptShapeKind.Group:
                    OfficeXmlNode tree = slide.PartRef.Root.Path(LegacyPptXml.P, "cSld", "spTree");
                    if (tree.IsNone) throw new InvalidDataException("The projected slide has no shape tree.");
                    uint nextShapeId = tree.Descendants(LegacyPptXml.P, "cNvPr")
                        .Select(item => item.UInt("id") ?? 0U)
                        .DefaultIfEmpty(1U)
                        .Max() + 1U;
                    OfficeXmlNode group = CreateLegacyOpenXmlShape(slide.PartRef, tree, shape,
                        ref nextShapeId, slidePartsByLegacyId, soundContext,
                        projectedShapeIds);
                    if (!group.IsNone) {
                        slide.ReserveShapeIdsThrough(nextShapeId);
                        slide.PartRef.MarkEdited();
                    }
                    return group;
            }
            if (projectedShape != null) projectedElement = projectedShape.XmlElement;
            if (projectedElement.IsNone) return default;
            OfficeXmlNode projectedProperties = projectedElement.NamespaceId == LegacyPptXml.P
                && projectedElement.LocalName is "sp" or "cxnSp" or "pic"
                ? projectedElement.Child(LegacyPptXml.P, "spPr")
                : default;
            if (!projectedProperties.IsNone) {
                if (shape.Kind != LegacyPptShapeKind.Picture) {
                    ApplyLegacyShapeStyle(projectedProperties, shape);
                }
                ApplyLegacyShapeTransform(projectedProperties.GetOrAddChild(LegacyPptXml.A, "xfrm",
                    LegacyPptImportXml.ShapePropertiesOrder), shape);
                OfficeXmlNode preset = projectedProperties.Child(LegacyPptXml.A, "prstGeom");
                if (!preset.IsNone) {
                    LegacyPptShapeGeometryMapper.ApplyExactPresetAdjustments(shape.OfficeArtShapeType,
                        shape.Geometry, preset);
                }
            }
            if (projectedElement.Is(LegacyPptXml.P, "pic")) {
                ApplyLegacyPictureEditingProperties(projectedElement, shape);
                OfficeXmlNode blipFill = projectedElement.Child(LegacyPptXml.P, "blipFill");
                ApplyLegacyPictureCrop(blipFill, shape);
                ApplyLegacyPictureEffects(blipFill.Child(LegacyPptXml.A, "blip"), shape);
            }
            ApplyLegacyPlaceholder(projectedElement, shape);
            ApplyLegacyShapeMetadata(projectedElement, shape);
            ApplyLegacyShapeInteractions(slide.PartRef, projectedElement, shape,
                slidePartsByLegacyId, soundContext);
            RegisterLegacyShapeId(shape, projectedElement,
                projectedShapeIds);
            slide.PartRef.MarkEdited();
            return projectedElement;
        }

        private static void RegisterLegacyShapeId(LegacyPptShape source,
            OfficeXmlNode projected,
            IDictionary<uint, uint>? projectedShapeIds) {
            if (projectedShapeIds == null) return;
            uint? projectedId = GetLegacyNonVisualDrawingProperties(projected).UInt("id");
            if (projectedId.HasValue) projectedShapeIds[source.ShapeId] = projectedId.Value;
        }

        private static void ProjectLegacyConnectorRules(PowerPointSlide slide,
            IReadOnlyList<LegacyPptConnectorRule> rules,
            IReadOnlyDictionary<uint, uint> projectedShapeIds) {
            if (rules.Count == 0) return;
            OfficeXmlNode tree = slide.PartRef.Root.Path(LegacyPptXml.P, "cSld", "spTree");
            if (tree.IsNone) return;
            ProjectLegacyConnectorRules(tree, rules, projectedShapeIds);
            slide.PartRef.MarkEdited();
        }

        private static readonly OfficeXmlChildOrder LegacyConnectionShapeOrder = new(
            new[] { LegacyPptXml.P, LegacyPptXml.P, LegacyPptXml.P, LegacyPptXml.P },
            new[] { "nvCxnSpPr", "spPr", "style", "extLst" });

        private static void ProjectLegacyConnectorRules(OfficeXmlNode tree,
            IReadOnlyList<LegacyPptConnectorRule> rules,
            IReadOnlyDictionary<uint, uint> projectedShapeIds) {
            var connectorsById = new Dictionary<uint, OfficeXmlNode>();
            foreach (OfficeXmlNode connector in tree.Descendants(LegacyPptXml.P, "cxnSp")) {
                uint? id = connector.Path(LegacyPptXml.P, "nvCxnSpPr", "cNvPr").UInt("id");
                if (id.HasValue && !connectorsById.ContainsKey(id.Value)) {
                    connectorsById.Add(id.Value, connector);
                }
            }
            foreach (LegacyPptConnectorRule rule in rules) {
                if (!projectedShapeIds.TryGetValue(rule.ConnectorShapeId, out uint connectorId)) continue;
                if (!connectorsById.TryGetValue(connectorId,
                        out OfficeXmlNode connector)) continue;
                OfficeXmlNode nonVisual = connector.GetOrAddChild(LegacyPptXml.P, "nvCxnSpPr", LegacyConnectionShapeOrder);
                OfficeXmlNode drawingProperties = nonVisual.GetOrAddChild(LegacyPptXml.P, "cNvCxnSpPr",
                    LegacyPptImportXml.NonVisualConnectorOrder);
                drawingProperties.RemoveChildren(LegacyPptXml.A, "stCxn");
                drawingProperties.RemoveChildren(LegacyPptXml.A, "endCxn");
                if (projectedShapeIds.TryGetValue(rule.StartShapeId, out uint startShapeId)) {
                    drawingProperties.AddA("stCxn")
                        .Set("id", (long)startShapeId)
                        .Set("idx", (long)rule.StartConnectionSiteIndex);
                }
                if (projectedShapeIds.TryGetValue(rule.EndShapeId, out uint endShapeId)) {
                    drawingProperties.AddA("endCxn")
                        .Set("id", (long)endShapeId)
                        .Set("idx", (long)rule.EndConnectionSiteIndex);
                }
            }
        }

        private static void ApplyLegacyOlePreview(PowerPointPartRef slidePart,
            OfficeXmlNode target, LegacyPptShape source) {
            if (source.Picture?.HasImportableImage != true
                || source.Picture.ContentType == null
                || !target.Is(LegacyPptXml.P, "graphicFrame")) return;
            OfficeXmlNode ole = target.Path(LegacyPptXml.A, "graphic", "graphicData")
                .Child(LegacyPptXml.P, "oleObj");
            OfficeXmlNode picture = ole.Child(LegacyPptXml.P, "pic");
            if (picture.IsNone) return;
            string relationshipId = AddLegacyImagePart(slidePart,
                source.Picture);
            OfficeXmlNode blipFill = picture.SetChild(LegacyPptXml.P, "blipFill", LegacyPptImportXml.PictureOrder);
            OfficeXmlNode blip = blipFill.AddA("blip").SetRelationship("embed", relationshipId);
            blipFill.AddA("stretch").AddA("fillRect");
            ApplyLegacyPictureCrop(blipFill, source);
            ApplyLegacyPictureEffects(blip, source);
            slidePart.MarkEdited();
        }

        /// <summary>A placeholder for a table on a master or special master, replaced once the slides are projected.</summary>
        private static OfficeXmlNode AddLegacyLateTable(LegacyPptImportParts parts, PowerPointPartRef ownerPart,
            OfficeXmlNode parent, LegacyPptShape source, uint shapeId) {
            OfficeXmlNode frame = parent.AddP("graphicFrame");
            OfficeXmlNode nonVisual = frame.AddP("nvGraphicFramePr");
            nonVisual.AddP("cNvPr").Set("id", (long)shapeId).Set("name", $"Binary Table {shapeId - 1U}");
            nonVisual.AddP("cNvGraphicFramePr");
            nonVisual.AddP("nvPr");
            parts.LateTables.Add(new LegacyPptLateTable(ownerPart, shapeId, source));
            return frame;
        }

        private static void ProjectLegacyLateTables(LegacyPptImportParts parts,
            IReadOnlyDictionary<uint, PowerPointPartRef> slidePartsByLegacyId,
            LegacyPptSoundProjectionContext soundContext) {
            if (parts.LateTables.Count == 0) return;
            EnsureLegacySdk(parts.Document);
            foreach (LegacyPptLateTable table in parts.LateTables) {
                OfficeXmlNode placeholder = FindLegacyShapeById(table.Owner.Root, table.ShapeId);
                if (placeholder.IsNone) {
                    throw new InvalidDataException("A projected master table lost its place before the slides were projected.");
                }
                OfficeXmlNode frame = placeholder.Parent.InsertBefore(CreateLegacyTableFrameElement(table.Owner,
                    table.Source, table.ShapeId, slidePartsByLegacyId, soundContext, deferredInteractions: null), placeholder);
                placeholder.Remove();
                ApplyLegacyShapeMetadata(frame, table.Source);
                ApplyLegacyShapeInteractions(table.Owner, frame, table.Source,
                    slidePartsByLegacyId, soundContext);
                table.Owner.MarkEdited();
            }
            parts.LateTables.Clear();
        }

        internal sealed class LegacyPptLateTable {
            internal LegacyPptLateTable(PowerPointPartRef owner, uint shapeId, LegacyPptShape source) {
                Owner = owner;
                ShapeId = shapeId;
                Source = source;
            }

            internal PowerPointPartRef Owner { get; }
            internal uint ShapeId { get; }
            internal LegacyPptShape Source { get; }
        }

        private static OfficeImageFormat GetLegacyPicturePartType(string contentType) => contentType switch {
            "image/png" => OfficeImageFormat.Png,
            "image/jpeg" => OfficeImageFormat.Jpeg,
            "image/bmp" => OfficeImageFormat.Bmp,
            "image/tiff" => OfficeImageFormat.Tiff,
            "image/x-emf" => OfficeImageFormat.Emf,
            "image/x-wmf" => OfficeImageFormat.Wmf,
            _ => throw new NotSupportedException($"Legacy picture content type '{contentType}' is not supported.")
        };

        private static PowerPointPlaceholderType? MapPlaceholder(LegacyPptPlaceholderKind placeholder) {
            switch (placeholder) {
                case LegacyPptPlaceholderKind.MasterTitle:
                case LegacyPptPlaceholderKind.Title: return PowerPointPlaceholderType.Title;
                case LegacyPptPlaceholderKind.MasterCenterTitle:
                case LegacyPptPlaceholderKind.CenterTitle: return PowerPointPlaceholderType.CenteredTitle;
                case LegacyPptPlaceholderKind.MasterSubtitle:
                case LegacyPptPlaceholderKind.Subtitle: return PowerPointPlaceholderType.SubTitle;
                case LegacyPptPlaceholderKind.MasterBody:
                case LegacyPptPlaceholderKind.Body: return PowerPointPlaceholderType.Body;
                case LegacyPptPlaceholderKind.VerticalTitle: return PowerPointPlaceholderType.Title;
                case LegacyPptPlaceholderKind.VerticalBody: return PowerPointPlaceholderType.Body;
                case LegacyPptPlaceholderKind.MasterNotesSlideImage:
                case LegacyPptPlaceholderKind.NotesSlideImage: return PowerPointPlaceholderType.SlideImage;
                case LegacyPptPlaceholderKind.MasterNotesBody:
                case LegacyPptPlaceholderKind.NotesBody: return PowerPointPlaceholderType.Body;
                case LegacyPptPlaceholderKind.MasterDate: return PowerPointPlaceholderType.DateAndTime;
                case LegacyPptPlaceholderKind.MasterSlideNumber: return PowerPointPlaceholderType.SlideNumber;
                case LegacyPptPlaceholderKind.MasterFooter: return PowerPointPlaceholderType.Footer;
                case LegacyPptPlaceholderKind.MasterHeader: return PowerPointPlaceholderType.Header;
                case LegacyPptPlaceholderKind.Graph: return PowerPointPlaceholderType.Chart;
                case LegacyPptPlaceholderKind.Table: return PowerPointPlaceholderType.Table;
                case LegacyPptPlaceholderKind.ClipArt: return PowerPointPlaceholderType.ClipArt;
                case LegacyPptPlaceholderKind.Media: return PowerPointPlaceholderType.Media;
                case LegacyPptPlaceholderKind.Picture: return PowerPointPlaceholderType.Picture;
                case LegacyPptPlaceholderKind.Object:
                case LegacyPptPlaceholderKind.VerticalObject: return PowerPointPlaceholderType.Object;
                case LegacyPptPlaceholderKind.OrganizationChart: return PowerPointPlaceholderType.Diagram;
                default: return null;
            }
        }

        private static long ToEmus(int masterUnits) => checked((long)Math.Round(
            masterUnits * EmusPerLegacyMasterUnit, MidpointRounding.AwayFromZero));
    }
}
