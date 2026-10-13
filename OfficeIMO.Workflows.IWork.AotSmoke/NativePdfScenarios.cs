using System.Globalization;
using System.Text.RegularExpressions;
using OfficeIMO;
using OfficeIMO.Pdf;
using OfficeIMO.Workflows;
using OfficeIMO.Workflows.IWork;
using static NativePdfAssets;

internal static class NativePdfScenarios {
    internal static async Task RunAsync(string root) {
        foreach (Source source in new[] { Body, List, Slides, UnpaddedControl }) {
            OfficeWorkflowRequest request = Request(root, source, "complete-" + Path.GetFileNameWithoutExtension(source.Name), PortableSettings());
            OfficeWorkflowResult result = await IWorkWorkflow.CreateRunner().RunAsync(request);
            Require(result.Succeeded, source.Name + ": " + result.Summary);
            VerifyEvidence(result, source);
            PdfReadDocument pdf = PdfReadDocument.Open(File.ReadAllBytes(request.OutputPath!));
            Require(pdf.Pages.Count == source.Pages, source.Name + " lost pages.");
            VerifyContent(pdf, source);
            VerifyIntegrity(request, source);
            Console.WriteLine($"PASS | {source.Name} | {source.Pages} PDF pages | {Hash(request.OutputPath!)} | complete stages, saved/reopened content and source preservation");
        }

        await RejectAsync(root, NativeNumbers, "native-padding", PortableSettings(), "padding");
        var missingFont = new IWorkWorkflowSettings();
        missingFont.PagesPdfOptions.ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic();
        missingFont.PagesPdfOptions.TextFallbacks = PdfTextFallbackFeatures.None;
        await RejectAsync(root, Body, "missing-font", missingFont, "font", requireStages: true);
        IWorkWorkflowSettings partial = PortableSettings(); partial.ConversionOptions.AllowPartialEditableReconstruction = true;
        await RejectAsync(root, Body, "partial-policy", partial, "complete editable");
        IWorkWorkflowSettings preview = PortableSettings(); preview.ConversionOptions.Mode = OfficeIMO.IWork.IWorkConversionMode.VisualOnly;
        await RejectAsync(root, Body, "preview-policy", preview, "complete editable");
        IWorkWorkflowSettings selection = PortableSettings(); selection.NumbersPdfOptions.MaxRowsPerSheet = 1;
        await RejectAsync(root, UnpaddedControl, "row-selection", selection, "every sheet, row, and column");
        IWorkWorkflowSettings exclusion = PortableSettings(); exclusion.KeynotePdfOptions.IncludePictures = false;
        await RejectAsync(root, Slides, "content-exclusion", exclusion, "content type");
        await RejectAsync(root, Body, "output-limit", PortableSettings(), "output limit", requireStages: true, outputLimit: 32);

        await CancelAsync(root, atPublication: false);
        await CancelAsync(root, atPublication: true);
    }

    private static OfficeWorkflowRequest Request(string root, Source source, string name, IWorkWorkflowSettings settings) {
        string directory = Path.Combine(root, name); Directory.CreateDirectory(directory);
        return new OfficeWorkflowRequest {
            InputPath = PathFor(source), OutputPath = Path.Combine(directory, "output.pdf"),
            Operation = OfficeWorkflowOperation.Convert, ConversionRouteId = source.Route,
            ConflictPolicy = OfficeWorkflowConflictPolicy.Replace, RegisteredConversionSettings = settings
        };
    }

    private static void VerifyEvidence(OfficeWorkflowResult result, Source source) {
        OfficeWorkflowConversionEvidence evidence = result.ConversionEvidence
            ?? throw new InvalidOperationException(source.Name + " lost conversion evidence.");
        Require(evidence.Facts["conversionStages"] == "iwork-editable,pdf"
            && evidence.Facts["partialEditableReconstruction"] == "False"
            && evidence.Facts["omittedSourceUnitCount"] == "0" && evidence.Facts["unassessedSourceUnitCount"] == "0"
            && evidence.Facts["sourceUnitCount"] == evidence.Facts["reconstructedSourceUnitCount"],
            source.Name + " lost complete source coverage or stage evidence.");
        Require(!evidence.FidelityDiagnostics.Any(item => item.LossKind is OfficeConversionLossKind.Omission or OfficeConversionLossKind.Failure),
            source.Name + " published a reported omission or failure.");
        Require(evidence.FidelityDiagnostics.Any(item => item.Code == "IWORK_RECORD_FIDELITY_UNASSESSED"),
            source.Name + " concealed unassessed record fidelity.");
        Require(result.Diagnostics.Any(item => item.Code == "OutputReopened")
            && result.Diagnostics.Single(item => item.Code == "SourceSnapshot").Details["sha256"].Equals(source.Hash, StringComparison.OrdinalIgnoreCase),
            source.Name + " lost output reopen or source identity evidence.");
    }

    private static void VerifyContent(PdfReadDocument pdf, Source source) {
        string[] pages = pdf.Pages.Select(page => page.ExtractText()).ToArray();
        if (source == Body) {
            Require(pages[0].Contains("OfficeIMO native body document", StringComparison.Ordinal)
                && pages[0].Contains("Bold 😀 plain edited — Zażółć café", StringComparison.Ordinal), "Body title or Unicode changed.");
            string text = string.Join("\n", pages); int previous = -1;
            for (int number = 1; number <= 80; number++) {
                string paragraph = "Paragraph " + number.ToString("00", CultureInfo.InvariantCulture)
                    + ": the body flows over pages with editable text.";
                int at = text.IndexOf(paragraph, StringComparison.Ordinal);
                Require(at > previous && text.IndexOf(paragraph, at + paragraph.Length, StringComparison.Ordinal) == -1,
                    "Body paragraph was omitted, duplicated or reordered."); previous = at;
            }
            Require(pages[^1].Contains("Paragraph 80:", StringComparison.Ordinal), "The final body page was lost.");
        } else if (source == List) {
            string text = Regex.Replace(pages[0], @"\s+", string.Empty);
            Require(text.Contains("7.Rootnativeserifmonoplain", StringComparison.Ordinal)
                && text.Contains("1.Deepnativeserifmonoplain2.Deeptwoserifmonoplain", StringComparison.Ordinal)
                && text.Contains("(3)Serifmarkerserifmonoplain", StringComparison.Ordinal)
                && text.Contains("5)Monomarkerserifmonoplain", StringComparison.Ordinal)
                && text.Contains("3.Nativechildplain", StringComparison.Ordinal), "Mixed-size list content or marker forms changed.");
            var spans = pdf.Pages[0].GetTextSpans();
            double[] sizes = [18, 12, 12, 12, 12, 12, 12, 12, 12, 12, 8, 8, 12, 12, 12, 12, 12, 8];
            Require(spans.Where(span => Regex.IsMatch(span.Text.Trim(), @"^(\d+\.|\(\d+\)|\d+\))$"))
                .Select(span => span.FontSize).SequenceEqual(sizes), "Mixed-size list marker sizes changed.");
        } else if (source == Slides) {
            string[] titles = ["Edited PNG — image behind text", "Edited alpha PNG", "Native checked JPEG"];
            for (int index = 0; index < pages.Length; index++) {
                Require(pages[index].Contains(titles[index], StringComparison.Ordinal), "A Keynote title was lost.");
                var placement = pdf.Pages[index].GetImagePlacements().Single();
                double height = index == 2 ? 240 : 320, top = index == 0 ? 110 : 120;
                Require(placement.X == (index == 2 ? 80 : 60) && placement.Width == (index == 2 ? 600 : 480)
                    && placement.Height == height && placement.Y == 540 - top - height, "Keynote image geometry changed.");
            }
            Require(pages[0].Contains("Covered source label", StringComparison.Ordinal) && pages[0].Contains("TEXT IN FRONT", StringComparison.Ordinal)
                && pages[1].Contains("100% alpha", StringComparison.Ordinal) && pages[2].Contains("JPEG / stretched frame", StringComparison.Ordinal),
                "Keynote text or stacking content changed.");
            var images = pdf.ExtractImages();
            Require(images.Count == 3 && images.All(image => image.Width == 96 && image.Height == 64)
                && images.Count(image => image.HasTransparencyMask && image.TransparencyMaskResolved) == 1,
                "Keynote raster data or transparency changed.");
        } else Require(pages[0].Contains("42", StringComparison.Ordinal), "The synthetic Numbers scalar was lost.");
    }

    private static async Task RejectAsync(string root, Source source, string name, IWorkWorkflowSettings settings,
        string? reason, bool requireStages = false, long? outputLimit = null) {
        OfficeWorkflowRequest request = Request(root, source, name, settings);
        File.WriteAllText(request.OutputPath!, "keep existing");
        if (outputLimit.HasValue) request.Limits.MaximumOutputBytes = outputLimit.Value;
        OfficeWorkflowResult result = await IWorkWorkflow.CreateRunner().RunAsync(request);
        Require(!result.Succeeded && File.ReadAllText(request.OutputPath!) == "keep existing", name + " published invalid or partial output.");
        if (reason != null) Require(result.Summary.Contains(reason, StringComparison.OrdinalIgnoreCase), name + ": " + result.Summary);
        if (outputLimit.HasValue) Require(result.FailureKind == OfficeWorkflowFailureKind.OutputFailed,
            name + " lost output-budget failure classification.");
        if (requireStages) Require(result.ConversionEvidence?.Facts["conversionStages"] == "iwork-editable,pdf", name + " lost stage evidence.");
        VerifyIntegrity(request, source);
        Console.WriteLine("PASS | rejection " + name + " | existing destination, source and staging preserved");
    }

    private static async Task CancelAsync(string root, bool atPublication) {
        OfficeWorkflowRequest request = Request(root, Body, atPublication ? "cancel-publication" : "cancel-start", PortableSettings());
        File.WriteAllText(request.OutputPath!, "keep existing");
        using var cancellation = new CancellationTokenSource(); bool reachedPublication = false;
        if (!atPublication) cancellation.Cancel();
        OfficeWorkflowResult result = await IWorkWorkflow.CreateRunner().RunAsync(request, new InlineProgress(update => {
            if (update.Stage != "publish") return;
            reachedPublication = true; cancellation.Cancel();
        }), cancellation.Token);
        Require(result.Status == OfficeWorkflowStatus.Cancelled && File.ReadAllText(request.OutputPath!) == "keep existing",
            "Cancellation replaced the destination or lost its status.");
        if (atPublication) Require(reachedPublication && result.Diagnostics.Any(item => item.Code == "OutputReopened"),
            "Publication cancellation did not reach a reopened completed PDF.");
        VerifyIntegrity(request, Body);
        Console.WriteLine("PASS | " + (atPublication ? "publication" : "pre-start") + " cancellation | destination and source preserved; staging removed");
    }

    private static void VerifyIntegrity(OfficeWorkflowRequest request, Source source) {
        Require(Hash(request.InputPath) == source.Hash, source.Name + " source changed.");
        Require(!Directory.EnumerateFiles(Path.GetDirectoryName(request.OutputPath!)!, "*.tmp", SearchOption.AllDirectories).Any(),
            source.Name + " leaked staging files.");
    }

    private sealed class InlineProgress(Action<OfficeWorkflowProgress> callback) : IProgress<OfficeWorkflowProgress> {
        public void Report(OfficeWorkflowProgress update) => callback(update);
    }
}
