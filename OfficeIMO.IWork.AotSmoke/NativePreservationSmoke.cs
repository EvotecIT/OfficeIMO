using System.Security.Cryptography;
using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.IWork;

internal static class NativePreservationSmoke {
    internal static void Run() {
        foreach (var (name, hash) in new[] {
            ("preservation.pages", "DC07FA4028AC9F261FDB7CB23F8B409CD80E092CDF8A8ED938ADC420B37AD6B4"),
            ("preservation.numbers", "C907C567584786CE9C5183A376F61D813F3BF56757DE531CBDBFB1FEFB4FCFD6"),
            ("preservation.key", "C0BC70026C2C977FDE0879F37849C14CD819DC055A3CD7F95E97CAC789828A17") }) {
            byte[] input = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, name));
            Require(Convert.ToHexString(SHA256.HashData(input)) == hash, "Native preservation fixture changed.");
            var source = IWorkSourceDocument.Open(input); var session = source.BeginEdit();
            var output = session.SaveBytesResult(); var reopened = IWorkSourceDocument.Open(output.GetPackageBytes());
            Require(output.Report.SourcePackageContentHash == output.Report.OutputPackageContentHash
                && output.Report.UpdatedRecords.Count == 0 && source.Records.Count == reopened.Records.Count,
                "Unchanged native save lost source content or record membership.");
            foreach (var entry in source.Entries)
                Require(entry.GetBytes().SequenceEqual(reopened.Entries.Single(item => item.Path == entry.Path).GetBytes()),
                    "Unchanged native save lost a file-entry payload.");
            using var destination = new MemoryStream(); destination.WriteByte(9); session.Save(destination);
            Require(destination.CanWrite && destination.ToArray().Skip(1).SequenceEqual(output.GetPackageBytes()),
                "Preservation save did not respect the caller's stream.");
            Console.WriteLine($"PASS | native unchanged {source.Kind} save | {hash} | {source.Records.Count} records | complete file-entry identity and caller stream");
        }
        var keynote = IWorkSourceDocument.Open(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "preservation.key")));
        var image = Jpeg(keynote); var edit = keynote.BeginEdit();
        edit.MoveKeynoteImage(image.SourceIdentity!, 100, 130);
        try { edit.SaveBytes(); throw new InvalidOperationException("Edited save accepted unassessed native visual caches implicitly."); }
        catch (InvalidOperationException ex) when (ex.Message.Contains("AllowUnassessedVisualCaches", StringComparison.Ordinal)) { }
        var translated = edit.SaveBytesResult(new() { AllowUnassessedVisualCaches = true });
        Require(Convert.ToHexString(SHA256.HashData(translated.GetPackageBytes())) == "D9F69DD9C3511CC34AA9B9766061308F29DC7871DB8F8F1FA23036A8122A800F",
            "NativeAOT image translation differs from the independently decoded Apple-accepted package.");
        var moved = Jpeg(IWorkSourceDocument.Open(translated.GetPackageBytes()));
        Require(moved.Geometry!.LeftPoints == 100 && moved.Geometry.TopPoints == 130 && moved.GetBytes().SequenceEqual(image.GetBytes())
            && image.Geometry!.LeftPoints == 80 && translated.Report.UpdatedRecords.Count == 1,
            "Image translation lost geometry, image bytes, immutable source or edit reporting.");
        using var conversion = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(new MemoryStream(translated.GetPackageBytes()),
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly });
        conversion.RequireCompleteEditableReconstruction(); using var saved = new MemoryStream(); conversion.Value.Save(saved); saved.Position = 0;
        using var presentation = PowerPointDocument.Load(saved);
        Require(presentation.Slides.Count == 3 && !presentation.ValidateDocument().Any(), "Translated source failed saved PowerPoint reconstruction.");
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        byte[] before = saved.ToArray();
        try { edit.Save(saved, cancellationToken: cancellation.Token); throw new InvalidOperationException("Native save ignored cancellation."); }
        catch (OperationCanceledException) { Require(before.SequenceEqual(saved.ToArray()), "Cancelled save changed the destination."); }
        Console.WriteLine("PASS | native Keynote image translation | D9F69DD9C3511CC34AA9B9766061308F29DC7871DB8F8F1FA23036A8122A800F | exact package, resources, cache acceptance, immutable source, strict saved/reopened PPTX and cancellation");
        byte[] tightInput = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "preservation.key"));
        var tight = IWorkSourceDocument.Open(tightInput, options: new() {
            MaximumPackageBytes = tightInput.Length, MaximumEntryBytes = 92_000, MaximumTotalEntryBytes = 92_000,
            MaximumIwaBytes = 32_768, MaximumDecompressedIwaBytes = 32_768, MaximumSnappyChunkBytes = 32_768,
            MaximumTotalDecompressedIwaBytes = 50_000, MaximumRecordCount = 205
        });
        Require(tight.BeginEdit().SaveBytesResult().Report.OutputPackageContentHash == tight.ComputePackageContentHash(),
            "Valid tight read bounds prevented unchanged preservation.");
        var tightEdit = tight.BeginEdit(); tightEdit.MoveKeynoteImage(Jpeg(tight).SourceIdentity!, 100, 130);
        Require(tightEdit.SaveBytes(new() { AllowUnassessedVisualCaches = true }).SequenceEqual(translated.GetPackageBytes()),
            "Valid tight read bounds prevented a bounded image translation or changed its bytes.");
        Console.WriteLine("PASS | native preservation tight source bounds | unchanged/translated saves, expanded carrier and aggregate entries, retained 32KiB decompression and 205-record limits");
    }

    private static IWorkImageAsset Jpeg(IWorkSourceDocument source) => source.ReadKeynote().Slides.SelectMany(slide => slide.Images)
        .Single(image => image.MediaType == "image/jpeg");
    private static void Require(bool condition, string message) { if (!condition) throw new InvalidOperationException(message); }
}
