using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json;
using OfficeIMO.Ocr;
using OfficeIMO.Ocr.Tesseract;

namespace OfficeIMO.PdfQualityCorpus;

// Opt-in native recognition measurements. Gold labels are verified before any OCR invocation.
internal static partial class OcrQualityCorpus {
    internal static async Task<int> RunAsync(string[] args) {
        if (args.Length is < 4 or > 5 || (args.Length == 5 && args[4] != "--require-quality"))
            throw new ArgumentException("Usage: ocr-quality <manifest.json> <asset-root> <output-directory> [--require-quality]");
        string manifestPath = Path.GetFullPath(args[1]), root = Path.GetFullPath(args[2]), output = Path.GetFullPath(args[3]);
        string reportPath = Path.Combine(output, "ocr-quality.json");
        if (File.Exists(reportPath)) throw new IOException("The OCR scorecard already exists; use a new output directory.");
        using var deadline = new CancellationTokenSource(TimeSpan.FromMinutes(15));
        ConsoleCancelEventHandler cancel = (_, eventArgs) => { eventArgs.Cancel = true; deadline.Cancel(); };
        Console.CancelKeyPress += cancel;
        try {
            var packet = await ReadManifestAsync(manifestPath, root, deadline.Token);
            Directory.CreateDirectory(output);
            TesseractOcrEngine provider = TesseractOcrEngine.CreateDefault();
            string version = await provider.GetVersionAsync(deadline.Token);
            string languages = string.Join("+", packet.Labels.SelectMany(label => label.Language.Split('+')).Distinct(StringComparer.Ordinal));
            TesseractLanguageDataResult models = await TesseractLanguageData.EnsureAsync(languages,
                new TesseractLanguageDataOptions { CacheDirectory = Path.Combine(output, "models") }, deadline.Token);
            var measurements = new List<OcrQualityMeasurement>();
            foreach (OcrQualityLabel label in packet.Labels) {
                OcrQualitySource source = packet.Manifest.Sources.Single(item => item.File == label.File);
                // Repetitions use identical source, raster preparation, labels, and provider configuration.
                for (int repetition = 1; repetition <= 2; repetition++) {
                    deadline.Token.ThrowIfCancellationRequested();
                    Stopwatch elapsed = Stopwatch.StartNew();
                    try {
                        byte[] bytes = await ReadVerifiedAsync(root, source.File, source.Sha256, deadline.Token);
                        var request = await PrepareRasterAsync(bytes, label, output, repetition, deadline.Token);
                        OcrResult? baseline = null;
                        var variants = new List<OcrRecognitionAttempt>();
                        foreach (int segmentation in new[] { 3, 6, 11 }) {
                            var engine = new TesseractOcrEngine(new TesseractOcrEngineOptions {
                                ExecutablePath = provider.ExecutablePath, TessdataDirectory = models.Directory,
                                Language = label.Language, PageSegmentationMode = segmentation, Dpi = 300,
                                Timeout = TimeSpan.FromSeconds(30), TemporaryDirectory = Path.Combine(output, "temporary")
                            });
                            variants.Add(new OcrRecognitionAttempt("psm-" + segmentation, new DelegateOcrEngine("tesseract-cli",
                                async (raster, token) => {
                                    OcrResult result = await engine.RecognizeAsync(raster, token);
                                    baseline ??= result;
                                    return result;
                                }, engine.Capabilities)));
                        }
                        var adaptive = new AdaptiveOcrEngine("tesseract-adaptive", variants, timeout: TimeSpan.FromSeconds(45));
                        AdaptiveOcrResult selected = await adaptive.RecognizeWithReviewAsync(request, deadline.Token);
                        ScanTextAccuracy before = ScanTextAccuracy.Measure(label.Expected, baseline!.Text, deadline.Token);
                        ScanTextAccuracy after = ScanTextAccuracy.Measure(label.Expected, selected.Result.Text, deadline.Token);
                        bool goldMet = GoldMet(label, after);
                        var thresholds = new List<OcrThresholdObservation>();
                        // Score the selected provider evidence before adaptive review diagnostics, which encode a different policy.
                        OcrResult evidence = selected.Result;
                        evidence.Diagnostics = evidence.Diagnostics.Where(item => !item.Code.StartsWith("adaptive-ocr-", StringComparison.Ordinal)).ToArray();
                        foreach (double confidence in new[] { .7, .8, .9, .95 }) {
                            OcrQualityAssessment assessment = new OcrReviewPolicy(confidence, .1).Assess(evidence);
                            thresholds.Add(new OcrThresholdObservation { MinimumWordConfidence = confidence,
                                MaximumUncertainWordFraction = .1, MeetsThresholds = assessment.MeetsThresholds,
                                GoldLimitsMet = goldMet, LowConfidenceWords = assessment.LowConfidenceWordCount,
                                UnknownConfidenceWords = assessment.UnknownConfidenceWordCount });
                        }
                        await File.WriteAllTextAsync(Path.Combine(output, label.Id + "-" + repetition + "-baseline.txt"), baseline.Text, deadline.Token);
                        await File.WriteAllTextAsync(Path.Combine(output, label.Id + "-" + repetition + "-selected.txt"), selected.Result.Text, deadline.Token);
                        bool geometry = GeometryWithinRaster(selected.Result, request);
                        bool unchanged = Hash(await ReadBoundedAsync(Resolve(root, source.File), deadline.Token)) == source.Sha256.ToUpperInvariant();
                        measurements.Add(new OcrQualityMeasurement { Id = label.Id, DocumentClass = label.DocumentClass,
                            SourceSha256 = source.Sha256, RasterSha256 = Hash(request.Payload), Repetition = repetition,
                            BaselineTextSha256 = Hash(System.Text.Encoding.UTF8.GetBytes(ScanTextAccuracy.Normalize(baseline.Text))),
                            SelectedTextSha256 = Hash(System.Text.Encoding.UTF8.GetBytes(ScanTextAccuracy.Normalize(selected.Result.Text))),
                            Baseline = before, Selected = after, GoldLimitsMet = goldMet,
                            ReviewRecommended = selected.ReviewRecommended, HasDisagreement = selected.HasDisagreement,
                            RetryIncomplete = selected.RetryIncomplete, GeometryWithinRaster = geometry, SourceUnchanged = unchanged,
                            SelectedAttempt = selected.SelectedAttempt, Attempts = selected.Attempts, Thresholds = thresholds,
                            ElapsedMilliseconds = elapsed.Elapsed.TotalMilliseconds,
                            HostPeakWorkingSetBytes = Process.GetCurrentProcess().PeakWorkingSet64 });
                        Console.WriteLine($"{label.Id} #{repetition}: CER {before.CharacterErrorRate:P2} -> {after.CharacterErrorRate:P2}; " +
                            $"attempts={selected.Attempts.Count}; review={selected.ReviewRecommended}; gold={goldMet}");
                    } catch (OperationCanceledException) when (deadline.IsCancellationRequested) { throw; }
                    catch (Exception error) {
                        measurements.Add(new OcrQualityMeasurement { Id = label.Id, DocumentClass = label.DocumentClass,
                            SourceSha256 = source.Sha256, Repetition = repetition, Outcome = "failed", FailureType = error.GetType().Name,
                            ElapsedMilliseconds = elapsed.Elapsed.TotalMilliseconds });
                        Console.WriteLine(label.Id + " #" + repetition + ": operational failure (" + error.GetType().Name + ")");
                    }
                }
            }
            bool operational = measurements.All(item => item.Outcome == "completed" && item.GeometryWithinRaster && item.SourceUnchanged);
            bool qualityMet = operational && measurements.All(item => item.GoldLimitsMet);
            var report = new {
                Schema = "officeimo.ocr.quality.v1", ManifestSha256 = packet.Hash, LabelsSha256 = packet.Manifest.LabelsSha256,
                Provider = provider.Id, ProviderVersion = version, Runtime = RuntimeInformation.FrameworkDescription,
                OperatingSystem = RuntimeInformation.OSDescription, Architecture = RuntimeInformation.ProcessArchitecture.ToString(),
                TrainedDataCatalogRevision = TesseractLanguageData.Version,
                LanguageModels = models.Files.Select(model => new { model.Language, model.Sha256, model.ByteCount }).ToArray(),
                Dpi = 300, SegmentationModes = new[] { 3, 6, 11 },
                Policy = new OcrReviewPolicy(), Repetitions = 2, OperationalQualificationPassed = operational, QualityLimitsMet = qualityMet,
                RepeatedOutputsStable = measurements.GroupBy(item => item.Id).All(group => group.Count() == 2 &&
                    group.All(item => item.Outcome == "completed") && group.Select(item => item.SelectedTextSha256).Distinct().Count() == 1 &&
                    group.Select(item => item.RasterSha256).Distinct().Count() == 1),
                Metric = "NFC with collapsed whitespace; case/punctuation retained; code-point CER and whitespace-token WER; rates may exceed one.",
                Limits = "Pinned finite corpus; no general accuracy claim. Threshold passage is not approval. Host peak is cumulative and excludes provider-process memory. Table-region text metrics do not qualify table structure. Hosted AI semantics are not measured.",
                ByDocumentClass = measurements.GroupBy(item => item.DocumentClass).Select(group => new {
                    DocumentClass = group.Key, Measurements = group.Count(), Completed = group.Count(item => item.Outcome == "completed"),
                    GoldPassed = group.Count(item => item.GoldLimitsMet), ReviewRecommended = group.Count(item => item.ReviewRecommended),
                    Improved = group.Count(item => item.Selected != null && item.Baseline != null && item.Selected.CharacterErrorRate < item.Baseline.CharacterErrorRate),
                    Regressed = group.Count(item => item.Selected != null && item.Baseline != null && item.Selected.CharacterErrorRate > item.Baseline.CharacterErrorRate),
                    BaselineCer = AggregateCer(group.Select(item => item.Baseline)), SelectedCer = AggregateCer(group.Select(item => item.Selected))
                }).ToArray(),
                ThresholdObservations = measurements.SelectMany(item => item.Thresholds).GroupBy(item => item.MinimumWordConfidence).Select(group => new {
                    MinimumWordConfidence = group.Key, Evaluated = group.Count(), ThresholdPassed = group.Count(item => item.MeetsThresholds),
                    ThresholdPassedWithGoldFailure = group.Count(item => item.MeetsThresholds && !item.GoldLimitsMet),
                    GoldPassedButReviewIndicated = group.Count(item => !item.MeetsThresholds && item.GoldLimitsMet)
                }).ToArray(), Cases = measurements
            };
            await PublishReportAsync(reportPath, JsonSerializer.SerializeToUtf8Bytes(report, Json), deadline.Token);
            return operational && (args.Length != 5 || qualityMet) ? 0 : 1;
        } catch (OperationCanceledException) when (deadline.IsCancellationRequested) {
            Console.Error.WriteLine("OCR qualification canceled; no final scorecard was published."); return 3;
        } finally { Console.CancelKeyPress -= cancel; }
    }

    private static bool GoldMet(OcrQualityLabel label, ScanTextAccuracy accuracy) =>
        accuracy.CharacterErrorRate <= label.MaximumCer && accuracy.WordErrorRate <= label.MaximumWer;

    private static double? AggregateCer(IEnumerable<ScanTextAccuracy?> measurements) {
        ScanTextAccuracy[] present = measurements.Where(item => item != null).Cast<ScanTextAccuracy>().ToArray();
        return present.Length == 0 ? null : (double)present.Sum(item => item.CharacterEdits) / Math.Max(1, present.Sum(item => item.ExpectedCharacters));
    }

    private static async Task PublishReportAsync(string path, byte[] bytes, CancellationToken token) {
        string temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try { await File.WriteAllBytesAsync(temporary, bytes, token); token.ThrowIfCancellationRequested(); File.Move(temporary, path); }
        finally { if (File.Exists(temporary)) File.Delete(temporary); }
    }
}
