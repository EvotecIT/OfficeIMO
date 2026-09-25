using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>Builds bounded, page-linked production evidence from one exact PDF snapshot.</summary>
internal static class PdfProductionPreflightInspector {
    internal static PdfProductionPreflightReport Inspect(PdfDocument source, PdfProductionPreflightOptions? options,
        CancellationToken cancellationToken) {
        Guard.NotNull(source, nameof(source));
        PdfProductionPreflightOptions effective = (options ?? new PdfProductionPreflightOptions()).Copy();
        effective.Validate();
        var snapshot = source.GetReadSnapshot(cancellationToken: cancellationToken);
        PdfReadDocument document = snapshot.Document;
        document.DemandContentExtraction("production preflight");
        PdfPageOptionalContentVisibility.DocumentState printVisibility =
            PdfPageOptionalContentVisibility.CreatePrintDocumentState(document.CatalogDictionary,
                document.Objects, document.ReadOptions.Limits.MaxContentNestingDepth,
                cancellationToken);
        int[] pageNumbers = effective.PageSelection?.ToPageNumbers(document.Pages.Count, nameof(effective.PageSelection))
            ?? Enumerable.Range(1, document.Pages.Count).ToArray();
        pageNumbers = pageNumbers.Distinct().ToArray();
        if (pageNumbers.Length > effective.MaxPages) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.RenderPages, effective.MaxPages, pageNumbers.Length);
        }

        var findings = new List<PdfProductionFinding>();
        var fixups = new List<PdfProductionFixupProposal>();
        InspectOutputIntents();
        PdfDocumentReadResult logical = PdfDocumentReadEngine.Read(document, new PdfReadOptions {
            Profile = PdfReadProfile.Fast,
            PageSelection = PdfPageSelection.From(pageNumbers),
            Pipeline = new PdfUnderstandingPipelineOptions { MaxPages = effective.MaxPages }
        }, cancellationToken);
        foreach (int pageNumber in pageNumbers) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfReadPage page = document.Pages[pageNumber - 1];
            PdfLogicalPage logicalPage = logical.PagesBySourcePageNumber[pageNumber][0];
            bool unresolvedPrintResources = logicalPage.HasOptionalContentUsage &&
                (printVisibility.HiddenObjectNumbers.Count > 0 || printVisibility.HasUnsupportedViewUsageApplications);
            InspectBoxes(pageNumber, page.GetGeometry());
            InspectFonts(pageNumber, unresolvedPrintResources);
            PdfReadPage printPage = page.WithOptionalContentVisibility(printVisibility);
            PdfPrintProductionColorEvidence color = InspectColor(pageNumber, printPage, unresolvedPrintResources);
            InspectImages(pageNumber, printPage, logicalPage, color);
        }
        return new PdfProductionPreflightReport(snapshot.Bytes, snapshot.Options, effective, pageNumbers, findings, fixups);

        void AddFinding(PdfProductionFinding finding) {
            if (findings.Count >= effective.MaxFindings) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.UnderstandingArtifacts, effective.MaxFindings, findings.Count + 1);
            }
            findings.Add(finding);
        }

        void AddFixup(int pageNumber, PdfPageBoundaryBox box, PdfPageBox bounds, string reason) {
            if (effective.MaxFixupProposals == 0) return;
            if (fixups.Count >= effective.MaxFixupProposals) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.UnderstandingArtifacts, effective.MaxFixupProposals, fixups.Count + 1);
            }
            fixups.Add(new PdfProductionFixupProposal(fixups.Count, pageNumber, box, bounds, reason));
        }

        void InspectOutputIntents() {
            IReadOnlyList<PdfOutputIntentInfo> intents = document.OutputIntents;
            bool strict = effective.Profile != PdfProductionPreflightProfile.GeneralPrint;
            if (intents.Count == 0) {
                AddFinding(new PdfProductionFinding(PdfProductionFindingKind.MissingOutputIntent,
                    strict ? PdfProductionFindingSeverity.Error : PdfProductionFindingSeverity.Warning,
                    null, "No catalog output intent was found; select and embed a suitable print profile explicitly."));
                return;
            }
            bool valid = document.OutputIntentsAreComplete && intents.All(IsInspectablePrintProfile);
            if (strict) valid = valid && intents.Count == 1 &&
                string.Equals(intents[0].Subtype, "GTS_PDFX", StringComparison.Ordinal) &&
                !string.IsNullOrWhiteSpace(intents[0].OutputConditionIdentifier);
            if (strict) {
                valid = valid && intents[0].DestinationOutputProfileColorComponents == 4 &&
                    string.Equals(intents[0].DestinationOutputProfileColorSpace, "CMYK", StringComparison.Ordinal);
            }
            if (!valid) {
                AddFinding(new PdfProductionFinding(PdfProductionFindingKind.InvalidOutputIntent,
                    strict ? PdfProductionFindingSeverity.Error : PdfProductionFindingSeverity.Warning,
                    null, "Output intent entries lack an inspectable print ICC profile or do not meet this profile's PDF/X candidate rules."));
            }
        }

        void InspectBoxes(int pageNumber, PdfPageGeometry geometry) {
            if (PdfPrintProductionStructureInspector.HasValidProductionBoxes(geometry)) return;
            AddFinding(new PdfProductionFinding(PdfProductionFindingKind.InvalidPageBoxes,
                effective.Profile == PdfProductionPreflightProfile.GeneralPrint ? PdfProductionFindingSeverity.Warning : PdfProductionFindingSeverity.Error,
                pageNumber, "Page production boxes are missing or violate MediaBox, BleedBox, and TrimBox/ArtBox nesting."));
            PdfPageBox? media = geometry.MediaBox;
            if (media is null || geometry.TrimBox is not null && geometry.ArtBox is not null) return;
            PdfPageBox? boundary = geometry.TrimBox ?? geometry.ArtBox;
            if (boundary is null) {
                PdfPageBox trim = geometry.CropBox is PdfPageBox crop && Contains(media, crop) ? crop : media;
                AddFixup(pageNumber, PdfPageBoundaryBox.TrimBox, trim,
                    "Set a provisional trim boundary from CropBox or MediaBox; review the intended finished size.");
                boundary = trim;
            }
            if (geometry.BleedBox is null && Contains(media, boundary)) {
                AddFixup(pageNumber, PdfPageBoundaryBox.BleedBox, media,
                    "Set a provisional bleed boundary to MediaBox; this does not create bleed artwork.");
            }
        }

        void InspectFonts(int pageNumber, bool unresolvedPrintResources) {
            if (unresolvedPrintResources) {
                AddFinding(new PdfProductionFinding(PdfProductionFindingKind.UninspectableFont,
                    PdfProductionFindingSeverity.Indeterminate, pageNumber,
                    "Font usage could not be isolated from print-hidden optional content."));
                return;
            }
            PdfPrintProductionStructureEvidence structure = PdfPrintProductionStructureInspector.Inspect(document, pageNumber, cancellationToken);
            if (structure.UnembeddedFontResourceCount > 0) {
                AddFinding(new PdfProductionFinding(PdfProductionFindingKind.UnembeddedFont,
                    effective.Profile == PdfProductionPreflightProfile.GeneralPrint ? PdfProductionFindingSeverity.Warning : PdfProductionFindingSeverity.Error,
                    pageNumber, $"{structure.UnembeddedFontResourceCount} reachable font resource(s) lack an inspectable embedded program."));
            }
            if (structure.UninspectableFontResourceCount > 0) {
                AddFinding(new PdfProductionFinding(PdfProductionFindingKind.UninspectableFont,
                    PdfProductionFindingSeverity.Indeterminate, pageNumber,
                    $"{structure.UninspectableFontResourceCount} font selection context(s) could not be inspected completely."));
            }
        }

        PdfPrintProductionColorEvidence InspectColor(int pageNumber, PdfReadPage printPage, bool unresolvedPrintResources) {
            PdfPrintProductionColorEvidence color = PdfPrintProductionColorInspector.Inspect(document, pageNumber, cancellationToken);
            if (unresolvedPrintResources) {
                (bool definiteRgb, bool definiteTransparency) = printPage.GetDefiniteUnlayeredPrintColorUse(cancellationToken);
                if (definiteRgb) {
                    AddFinding(new PdfProductionFinding(PdfProductionFindingKind.DeviceRgbColor,
                        effective.Profile == PdfProductionPreflightProfile.GeneralPrint ? PdfProductionFindingSeverity.Warning : PdfProductionFindingSeverity.Error,
                        pageNumber, "Always-visible page content uses device RGB."));
                }
                if (definiteTransparency && effective.Profile != PdfProductionPreflightProfile.PdfX4Candidate) {
                    AddFinding(new PdfProductionFinding(PdfProductionFindingKind.Transparency,
                        effective.Profile == PdfProductionPreflightProfile.PdfX1aCandidate ? PdfProductionFindingSeverity.Error : PdfProductionFindingSeverity.Warning,
                        pageNumber, "Always-visible page content uses transparency."));
                }
                AddFinding(new PdfProductionFinding(PdfProductionFindingKind.UninspectableColor,
                    PdfProductionFindingSeverity.Indeterminate, pageNumber,
                    "Color usage could not be isolated from print-hidden optional content."));
                return color;
            }
            if (color.HasDeviceRgbUsage) {
                AddFinding(new PdfProductionFinding(PdfProductionFindingKind.DeviceRgbColor,
                    effective.Profile == PdfProductionPreflightProfile.GeneralPrint ? PdfProductionFindingSeverity.Warning : PdfProductionFindingSeverity.Error,
                    pageNumber, "Reachable content uses device RGB without an explicit color conversion in this artifact."));
            }
            if (color.HasDeviceIndependentColorUsage && effective.Profile == PdfProductionPreflightProfile.PdfX1aCandidate) {
                AddFinding(new PdfProductionFinding(PdfProductionFindingKind.DeviceIndependentColor,
                    PdfProductionFindingSeverity.Error, pageNumber,
                    "Reachable page content uses device-independent color, which this PDF/X-1a candidate profile rejects."));
            }
            if (color.HasTransparency && effective.Profile != PdfProductionPreflightProfile.PdfX4Candidate) {
                AddFinding(new PdfProductionFinding(PdfProductionFindingKind.Transparency,
                    effective.Profile == PdfProductionPreflightProfile.PdfX1aCandidate ? PdfProductionFindingSeverity.Error : PdfProductionFindingSeverity.Warning,
                    pageNumber, "Reachable content uses transparency; review flattening and blending before press output."));
            }
            if (!color.IsComplete) {
                AddFinding(new PdfProductionFinding(PdfProductionFindingKind.UninspectableColor,
                    PdfProductionFindingSeverity.Indeterminate, pageNumber,
                    $"{color.UninspectableContentStreamCount} color or annotation appearance context(s) could not be inspected completely."));
            }
            return color;
        }

        void InspectImages(int pageNumber, PdfReadPage readPage, PdfLogicalPage logicalPage,
            PdfPrintProductionColorEvidence color) {
            if (logicalPage.HasOptionalContentUsage && printVisibility.HasUnsupportedViewUsageApplications) {
                AddFinding(new PdfProductionFinding(PdfProductionFindingKind.UninspectableImageResolution,
                    PdfProductionFindingSeverity.Indeterminate, pageNumber,
                    "The optional-content print configuration could not be evaluated completely."));
                if (color.HasUninspectedImagePlacementSources) {
                    AddFinding(new PdfProductionFinding(PdfProductionFindingKind.UninspectableImageResolution,
                        PdfProductionFindingSeverity.Indeterminate, pageNumber,
                        "A reachable tiling pattern or printable annotation may paint images whose effective resolution was not inspected."));
                }
                return;
            }
            IReadOnlyList<PdfImagePlacement> placements = readPage.GetImagePlacements(pageNumber, cancellationToken);
            int understoodPlacements = 0;
            var matchedPlacements = new HashSet<PdfImagePlacement>();
            foreach (PdfExtractedImage image in readPage.GetImages(pageNumber, placements, cancellationToken)) {
                foreach (PdfImagePlacement placement in PdfLogicalPage.MatchImagePlacements(image, placements)) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!matchedPlacements.Add(placement)) continue;
                    understoodPlacements++;
                    double userUnit = logicalPage.UserUnit ?? 1D;
                    double horizontalPoints = Math.Sqrt(placement.A * placement.A + placement.B * placement.B) * userUnit;
                    double verticalPoints = Math.Sqrt(placement.C * placement.C + placement.D * placement.D) * userUnit;
                    if (!IsPositiveFinite(horizontalPoints) || !IsPositiveFinite(verticalPoints) ||
                        image.Width <= 0 || image.Height <= 0) {
                        AddFinding(new PdfProductionFinding(PdfProductionFindingKind.UninspectableImageResolution,
                            PdfProductionFindingSeverity.Indeterminate, pageNumber, "A placed image lacks usable pixel or transform dimensions."));
                        continue;
                    }
                    double ppi = Math.Min(image.Width * 72D / horizontalPoints, image.Height * 72D / verticalPoints);
                    if (ppi >= effective.EffectiveMinimumImagePpi) continue;
                    PdfVisualBounds visual = logicalPage.TransformBoundsToVisual(placement.X, placement.Y,
                        placement.X + placement.Width, placement.Y + placement.Height);
                    AddFinding(new PdfProductionFinding(PdfProductionFindingKind.LowImageResolution,
                        effective.Profile == PdfProductionPreflightProfile.GeneralPrint ? PdfProductionFindingSeverity.Warning : PdfProductionFindingSeverity.Error,
                        pageNumber, $"Image placement resolves to {ppi:0.#} ppi, below {effective.EffectiveMinimumImagePpi:0.#} ppi.",
                        new PdfLogicalVisualBounds(visual.Left, visual.Top, visual.Right, visual.Bottom), ppi));
                }
            }
            if (understoodPlacements < placements.Count) {
                AddFinding(new PdfProductionFinding(PdfProductionFindingKind.UninspectableImageResolution,
                    PdfProductionFindingSeverity.Indeterminate, pageNumber,
                    $"{placements.Count - understoodPlacements} image placement(s) lacked usable extraction metadata."));
            }
            if (color.HasUninspectedImagePlacementSources) {
                AddFinding(new PdfProductionFinding(PdfProductionFindingKind.UninspectableImageResolution,
                    PdfProductionFindingSeverity.Indeterminate, pageNumber,
                    "A reachable tiling pattern or printable annotation may paint images whose effective resolution was not inspected."));
            }
        }

        bool IsInspectablePrintProfile(PdfOutputIntentInfo intent) {
            if (!intent.HasDestinationOutputProfile) return false;
            try {
                return intent.DestinationOutputProfileSizeBytes is int size && size >= 128 &&
                    intent.DestinationOutputProfileDeclaredSizeBytes == size &&
                    intent.DestinationOutputProfileHasIccSignature == true &&
                    string.Equals(intent.DestinationOutputProfileDeviceClass, "prtr", StringComparison.Ordinal) &&
                    intent.DestinationOutputProfileColorComponents == (intent.DestinationOutputProfileColorSpace switch {
                        "RGB " => 3,
                        "CMYK" => 4,
                        "GRAY" => 1,
                        _ => -1
                    }) &&
                    intent.DestinationOutputProfileHasSupportedOutputTransform == true;
            } catch (InvalidDataException) {
                return false;
            }
        }
    }

    private static bool Contains(PdfPageBox outer, PdfPageBox inner) =>
        inner.Left >= outer.Left && inner.Bottom >= outer.Bottom && inner.Right <= outer.Right && inner.Top <= outer.Top;

    private static bool IsPositiveFinite(double value) => value > 0D && !double.IsNaN(value) && !double.IsInfinity(value);
}
