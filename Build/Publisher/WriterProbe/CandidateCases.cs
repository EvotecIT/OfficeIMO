using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Maps semantic CLI cases to the single creation specification consumed by all codecs.</summary>
internal static class CandidateCases {
    private static readonly PublisherNativeCreationPlacement DefaultPlacement = new(72, 72, 288, 108);

    internal static PublisherNativeCreationPlan Create(string kind, string? value, string? secondValue = null) {
        if (kind == "sized-composition") return CandidateSizedComposition.Create(value, ReadImage(secondValue));
        if (kind is "multi-paragraph" or "multi-paragraph-linked" or "multi-paragraph-linked-hyphenated") {
            if (value != null || secondValue != null) throw new ArgumentException("The multi-paragraph specimens use fixed styled text.");
            return CandidateMultiParagraph.Create(kind != "multi-paragraph", kind == "multi-paragraph-linked-hyphenated");
        }
        if (kind is "caption-typing" or "caption-typing-language") {
            if (secondValue != null) throw new ArgumentException("The caption specimen accepts one text value.");
            return CandidateCaptionTyping.Create(value, kind == "caption-typing-language" ? 1033U : null);
        }
        if (kind is "multiple-masters" or "paired-masters") {
            if (value != null || secondValue != null) throw new ArgumentException("The master specimens use fixed text and assignments.");
            return CandidateMasterPages.Create(kind == "paired-masters");
        }
        if (kind is "frame-layout-language" or "linked-story-language" or "linked-story-language-overflow") {
            if (value != null || secondValue != null) throw new ArgumentException("The language specimens use fixed English text.");
            return kind == "frame-layout-language" ? CandidateFrameLayout.Create(false, 1033)
                : CandidateLinkedStory.Create(kind == "linked-story-language-overflow", 1033);
        }
        if (kind is "frame-layout" or "frame-layout-edited") {
            if (value != null || secondValue != null) throw new ArgumentException("The frame layout specimens use fixed text.");
            return CandidateFrameLayout.Create(kind == "frame-layout-edited");
        }
        if (kind is "table-cluster-near" or "table-cluster-near-edited" or "table-cluster-wide") {
            if (value != null || secondValue != null) throw new ArgumentException("The dense table specimens use fixed grids, text and paint.");
            return CandidateTableCluster.Create(kind == "table-cluster-wide" ? 20 : 18, kind == "table-cluster-near-edited");
        }
        if (kind is "table" or "table-edited") {
            if (value != null || secondValue != null) throw new ArgumentException("The table specimens use fixed cell text.");
            return CandidateTableCreation.Create(kind == "table-edited");
        }
        if (kind is "table-decorated" or "table-decorated-edited") {
            if (value != null || secondValue != null) throw new ArgumentException("The decorated table specimens use fixed cell text and paint.");
            return CandidateTableCreation.CreateDecorated(kind == "table-decorated-edited");
        }
        if (kind == "table-merged") {
            if (value != null || secondValue != null) throw new ArgumentException("The merged table specimen uses fixed cell text and spans.");
            return CandidateTableCreation.CreateMerged();
        }
        if (kind == "table-merged-decorated") {
            if (value != null || secondValue != null) throw new ArgumentException("The decorated merged table specimen uses fixed cell text, spans and paint.");
            return CandidateTableCreation.CreateMergedDecorated();
        }
        if (kind is "linked-story" or "linked-story-overflow") {
            if (value != null || secondValue != null) throw new ArgumentException("The linked story specimens use fixed text.");
            return CandidateLinkedStory.Create(kind == "linked-story-overflow");
        }
        if (kind is "table-grid" or "table-grid-empty" or "table-grid-edited") {
            if (secondValue != null) throw new ArgumentException("The grid specimen accepts one rows-by-columns value.");
            return kind == "table-grid-edited" ? CandidateTableGrid.CreateEdited(value) : CandidateTableGrid.Create(value, kind == "table-grid-empty");
        }
        if (kind == "master-content") {
            if (secondValue != null) throw new ArgumentException("The master content specimen accepts one image path.");
            return CandidateMasterContent.Create(ReadImage(value));
        }
        if (kind == "newsletter") {
            if (secondValue != null) throw new ArgumentException("The newsletter specimen accepts one image path.");
            return CandidateNewsletter.Create(ReadImage(value));
        }
        if (kind == "paired-brochure") {
            if (secondValue != null) throw new ArgumentException("The paired brochure specimen accepts one image path.");
            return CandidatePairedBrochure.Create(ReadImage(value));
        }
        if (kind == "mixed-object-pages") {
            if (secondValue != null) throw new ArgumentException("The mixed object specimen accepts one image path.");
            return CandidateMixedPages.Create(ReadImage(value));
        }
        if (kind == "shared-image-pages") return CandidateImagePages.Create(ReadImage(value), ReadImage(secondValue),
            includeNativeNames: true, shareFirstImage: true);
        if (kind == "cropped-image-pages") return CandidateImagePages.Create(ReadImage(value), ReadImage(secondValue),
            includeNativeNames: true, crop: new OfficeIMO.Drawing.OfficeImageSourceCrop(0.25, 0.125, 0.0625, 0.1875));
        if (kind is "image-pages" or "named-image-pages")
            return CandidateImagePages.Create(ReadImage(value), ReadImage(secondValue), includeNativeNames: kind == "named-image-pages");
        if (secondValue != null) throw new ArgumentException("Only image page specimens accept a second image path.");
        if (kind == "text-pages") return CandidateTextPages.Create(value);
        if (kind == "object-pages") return CandidateObjectPages.Create(ReadImage(value));
        if (kind == "blank-pages") {
            if (!int.TryParse(value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out int count)
                || count < 1 || count > 16) throw new ArgumentException("The blank-pages case requires a page count from 1 to 16.");
            return new PublisherNativeCreationPlan(612 * 12700, 792 * 12700, Array.Empty<PublisherNativeCreationObject>(), count);
        }
        if (kind is not ("blank" or "text" or "styled-text" or "shape" or "image" or "combined")) throw new ArgumentException("Unknown candidate kind.");
        if (kind == "styled-text" && value != null) throw new ArgumentException("The styled-text case uses its fixed three-run sample.");
        PublisherNativeCreationObject[] objects = kind switch {
            "blank" => Array.Empty<PublisherNativeCreationObject>(),
            "text" => new[] { new PublisherNativeCreationObject(PublisherNativeCreationObjectKind.TextBox, DefaultPlacement,
                PublisherNativeCreationStory.Plain(value ?? "Template-free Publisher writer probe\r")) },
            "styled-text" => new[] { new PublisherNativeCreationObject(PublisherNativeCreationObjectKind.TextBox, DefaultPlacement,
                PublisherNativeCreationStory.Styled(CandidateStyledText.Create())) },
            "shape" => new[] { new PublisherNativeCreationObject(PublisherNativeCreationObjectKind.Rectangle, DefaultPlacement, PublisherNativeCreationStory.Plain("\r")) },
            "combined" => new[] {
                new PublisherNativeCreationObject(PublisherNativeCreationObjectKind.TextBox, DefaultPlacement, PublisherNativeCreationStory.Styled(CandidateStyledText.Create())),
                new PublisherNativeCreationObject(PublisherNativeCreationObjectKind.Rectangle, new PublisherNativeCreationPlacement(72, 216, 144, 72), PublisherNativeCreationStory.Plain("\r")),
                new PublisherNativeCreationObject(PublisherNativeCreationObjectKind.Picture, new PublisherNativeCreationPlacement(330, 216, 96, 96), Image: ReadImage(value))
            },
            _ => new[] { new PublisherNativeCreationObject(PublisherNativeCreationObjectKind.Picture, DefaultPlacement, Image: ReadImage(value)) }
        };
        return new PublisherNativeCreationPlan(612 * 12700, 792 * 12700, objects);
    }

    private static byte[] ReadImage(string? path) {
        if (path == null) throw new ArgumentException("The image candidate requires a PNG or JPEG path.");
        if (new FileInfo(path).Length > 1024 * 1024) throw new ArgumentException("Probe images are limited to 1 MiB.");
        return File.ReadAllBytes(path);
    }
}
