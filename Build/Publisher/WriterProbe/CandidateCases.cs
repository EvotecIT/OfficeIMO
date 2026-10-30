namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Maps semantic CLI cases to the single creation specification consumed by all codecs.</summary>
internal static class CandidateCases {
    private static readonly CandidatePlacement DefaultPlacement = new(72, 72, 288, 108);

    internal static CandidatePublicationSpec Create(string kind, string? value) {
        if (kind is not ("blank" or "text" or "styled-text" or "shape" or "image")) throw new ArgumentException("Unknown candidate kind.");
        if (kind == "styled-text" && value != null) throw new ArgumentException("The styled-text case uses its fixed three-run sample.");
        CandidateObjectSpec[] objects = kind switch {
            "blank" => Array.Empty<CandidateObjectSpec>(),
            "text" => new[] { new CandidateObjectSpec(CandidateObjectKind.TextBox, DefaultPlacement,
                CandidateStorySpec.Plain(value ?? "Template-free Publisher writer probe\r")) },
            "styled-text" => new[] { new CandidateObjectSpec(CandidateObjectKind.TextBox, DefaultPlacement,
                CandidateStorySpec.Styled(CandidateStyledText.Create())) },
            "shape" => new[] { new CandidateObjectSpec(CandidateObjectKind.Rectangle, DefaultPlacement, CandidateStorySpec.Plain("\r")) },
            _ => new[] { new CandidateObjectSpec(CandidateObjectKind.Picture, DefaultPlacement, Image: ReadImage(value)) }
        };
        return new CandidatePublicationSpec(612 * 12700, 792 * 12700, objects);
    }

    private static byte[] ReadImage(string? path) {
        if (path == null) throw new ArgumentException("The image candidate requires a PNG or JPEG path.");
        if (new FileInfo(path).Length > 1024 * 1024) throw new ArgumentException("Probe images are limited to 1 MiB.");
        return File.ReadAllBytes(path);
    }
}
