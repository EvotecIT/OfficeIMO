using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public class PowerPointAdvancedFeatures {
        [Fact]
        public void CanHandleBackgroundFormattingTransitionsAndCharts() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");

            using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTextBox text = slide.AddTextBox("Test");
                slide.AddImage(imagePath);
                slide.AddTable(2,2);
                slide.AddChart();
                slide.Notes.Text = "Notes";

                slide.BackgroundColor = OfficeColor.Parse("FF0000");
                text.FillColor = OfficeColor.Parse("00FF00");
                slide.Transition = PowerPointSlideTransition.Fade;

                presentation.Save();
            }

            using (PowerPointDocument presentation = PowerPointDocument.Load(filePath)) {
                PowerPointSlide slide = presentation.Slides.Single();
                Assert.Equal(OfficeColor.Parse("FF0000"), slide.BackgroundColor);
                Assert.Equal(PowerPointSlideTransition.Fade, slide.Transition);
                Assert.Single(slide.TextBoxes);
                Assert.Single(slide.Images);
                Assert.Single(slide.Tables);
                Assert.Single(slide.Charts);
                Assert.Equal(OfficeColor.Parse("00FF00"), slide.TextBoxes.First().FillColor);
                Assert.Equal("Notes", slide.Notes.Text);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void MorphTransition_ValidatesAndRoundTrips() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    slide.Transition = PowerPointSlideTransition.Morph;
                    Assert.False(presentation.IsSdkLoaded);

                    presentation.Save();
                    Assert.Empty(presentation.ValidateDocument());
                }

                using (PowerPointDocument presentation = PowerPointDocument.Load(filePath)) {
                    PowerPointSlide slide = presentation.Slides.Single();
                    Assert.Equal(PowerPointSlideTransition.Morph, slide.Transition);
                    Assert.False(presentation.IsSdkLoaded);
                    Assert.Empty(presentation.ValidateDocument());
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.Single();
                    OpenXmlValidator validator = new(FileFormatVersions.Microsoft365);
                    List<ValidationErrorInfo> errors = validator.Validate(document).ToList();

                    Assert.Empty(errors);
                    Assert.Contains("AlternateContent", slidePart.Slide.OuterXml, StringComparison.Ordinal);
                    Assert.Contains("morph", slidePart.Slide.OuterXml, StringComparison.Ordinal);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void MorphTransitionKeepsTypedBranchesOnAnOpenSdkSlide() {
            using PowerPointDocument presentation = PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            var package = presentation.OpenXmlDocument;

            slide.Transition = PowerPointSlideTransition.Morph;

            Transition[] transitions = GetAlternateTransitions(package.PresentationPart!.SlideParts.Single());
            Assert.Equal(2, transitions.Length);
            Assert.IsType<FadeTransition>(transitions[1].FirstChild);
        }

        [Fact]
        public void TransitionIgnoresEffectMarkupInSlideExtensions() {
            using var packageStream = new MemoryStream();
            using (PowerPointDocument source = PowerPointDocument.Create()) {
                source.AddSlide();
                source.Save(packageStream);
            }

            packageStream.Position = 0;
            using (PresentationDocument package = PresentationDocument.Open(packageStream, true)) {
                const string slideXml = """
                    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
                      <p:cSld><p:spTree>
                        <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
                        <p:grpSpPr/>
                      </p:spTree></p:cSld>
                      <p:extLst><p:ext uri="{35079707-908B-43A0-BC40-3D0B4EA5C100}">
                        <p:transition advClick="1"><p:fade/></p:transition>
                      </p:ext></p:extLst>
                    </p:sld>
                    """;
                using var xml = new MemoryStream(Encoding.UTF8.GetBytes(slideXml), writable: false);
                package.PresentationPart!.SlideParts.Single().FeedData(xml);
            }

            packageStream.Position = 0;
            using PowerPointDocument presentation = PowerPointDocument.Load(packageStream);
            PowerPointSlide slide = Assert.Single(presentation.Slides);
            Assert.Equal(PowerPointSlideTransition.None, slide.Transition);
            Assert.Null(slide.TransitionAdvanceOnClick);

            slide.Transition = PowerPointSlideTransition.WipeDown;
            Assert.Equal(PowerPointSlideTransition.WipeDown, slide.Transition);
            Assert.False(presentation.IsSdkLoaded);
        }

        [Fact]
        public void MorphTransition_SoundActionsRemainOnMorphMarkup() {
            string filePath = Path.Combine(Path.GetTempPath(),
                Guid.NewGuid() + ".pptx");
            byte[] audio = { 82, 73, 70, 70, 4, 0, 0, 0, 87, 65, 86, 69 };

            try {
                using (PowerPointDocument presentation =
                       PowerPointDocument.Create(filePath)) {
                    PowerPointSlide start = presentation.AddSlide();
                    start.Transition = PowerPointSlideTransition.Fade;
                    start.TransitionSpeed = PowerPointSlideTransitionSpeed.Fast;
                    start.TransitionDurationSeconds = 0.8;
                    start.TransitionAdvanceOnClick = false;
                    start.TransitionAdvanceAfterSeconds = 2.5;
                    using (var sound = new MemoryStream(audio,
                               writable: false)) {
                        start.SetTransitionSound(sound, "Morph sound");
                    }
                    Assert.False(presentation.IsSdkLoaded);
                    start.Transition = PowerPointSlideTransition.Morph;

                    PowerPointSlide stop = presentation.AddSlide();
                    stop.Transition = PowerPointSlideTransition.Morph;
                    stop.StopTransitionSound();

                    Assert.Equal(PowerPointSlideTransition.Morph, start.Transition);
                    Assert.Equal(PowerPointSlideTransition.Morph, stop.Transition);
                    presentation.Save();
                }

                using PowerPointDocument reopened =
                    PowerPointDocument.Load(filePath);
                PowerPointSlide startSlide = reopened.Slides[0];
                PowerPointSlide stopSlide = reopened.Slides[1];
                Assert.Equal(PowerPointSlideTransition.Morph,
                    startSlide.Transition);
                Assert.True(startSlide.HasTransitionSound);
                Assert.Equal("Morph sound",
                    startSlide.TransitionSoundName);
                Assert.Equal(audio,
                    startSlide.GetTransitionSoundBytes());
                Assert.Equal(PowerPointSlideTransitionSpeed.Fast,
                    startSlide.TransitionSpeed);
                Assert.Equal(0.8,
                    startSlide.TransitionDurationSeconds);
                Assert.False(startSlide.TransitionAdvanceOnClick);
                Assert.Equal(2.5,
                    startSlide.TransitionAdvanceAfterSeconds);
                Assert.Equal(PowerPointSlideTransition.Morph,
                    stopSlide.Transition);
                Assert.True(stopSlide.TransitionStopsSound);
                Assert.False(reopened.IsSdkLoaded);
                Assert.Empty(reopened.ValidateDocument());

                using PresentationDocument document =
                    PresentationDocument.Open(filePath, false);
                SlidePart[] slideParts = document.PresentationPart!
                    .SlideParts.ToArray();
                Transition[] startTransitions = GetAlternateTransitions(
                    slideParts[0]);
                Assert.Equal(2, startTransitions.Length);
                Assert.All(startTransitions, transition => {
                    Assert.NotNull(transition.GetFirstChild<SoundAction>()?
                        .GetFirstChild<StartSoundAction>());
                    Assert.Equal(TransitionSpeedValues.Fast,
                        transition.Speed?.Value);
                    Assert.Equal("800", transition.Duration?.Value);
                    Assert.False(transition.AdvanceOnClick?.Value);
                    Assert.Equal("2500",
                        transition.AdvanceAfterTime?.Value);
                });
                Transition[] stopTransitions = GetAlternateTransitions(
                    slideParts[1]);
                Assert.Equal(2, stopTransitions.Length);
                Assert.All(stopTransitions, transition =>
                    Assert.NotNull(transition
                        .GetFirstChild<SoundAction>()?
                        .GetFirstChild<EndSoundAction>()));
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }

        private static Transition[] GetAlternateTransitions(
            SlidePart slidePart) {
            AlternateContent content = Assert.Single(slidePart.Slide
                .Elements<AlternateContent>());
            return content.Elements<AlternateContentChoice>()
                .Select(choice => choice.GetFirstChild<Transition>())
                .Concat(new[] { content
                    .GetFirstChild<AlternateContentFallback>()?
                    .GetFirstChild<Transition>() })
                .Where(transition => transition != null)
                .Cast<Transition>()
                .ToArray();
        }

        [Fact]
        public void AllSupportedTransitions_ValidateAndRoundTrip() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            PowerPointSlideTransition[] transitions = Enum.GetValues(typeof(PowerPointSlideTransition)).Cast<PowerPointSlideTransition>().ToArray();

            try {
                using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                    foreach (PowerPointSlideTransition transition in transitions) {
                        PowerPointSlide slide = presentation.AddSlide();
                        slide.AddTitle(transition.ToString());
                        slide.Transition = transition;
                    }

                    presentation.Save();
                    Assert.Empty(presentation.ValidateDocument());
                }

                using (PowerPointDocument presentation = PowerPointDocument.Load(filePath)) {
                    PowerPointSlide[] slides = presentation.Slides.ToArray();
                    Assert.Equal(transitions.Length, slides.Length);

                    for (int i = 0; i < transitions.Length; i++) {
                        Assert.Equal(transitions[i], slides[i].Transition);
                    }

                    Assert.Empty(presentation.InspectFeatures()
                        .FindFeatures("Unsupported transition markup"));
                    Assert.Empty(presentation.ValidateDocument());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void TransitionTimingAndAdvanceSettings_RoundTripAndValidatePackage() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    slide.Transition = PowerPointSlideTransition.Fade;
                    slide.TransitionSpeed = PowerPointSlideTransitionSpeed.Fast;
                    slide.TransitionDurationSeconds = 0.6;
                    slide.TransitionAdvanceOnClick = false;
                    slide.TransitionAdvanceAfterSeconds = 4.25;

                    presentation.Save();
                    Assert.Empty(presentation.ValidateDocument());
                }

                using (PowerPointDocument presentation = PowerPointDocument.Load(filePath)) {
                    PowerPointSlide slide = presentation.Slides.Single();
                    Assert.Equal(PowerPointSlideTransition.Fade, slide.Transition);
                    Assert.Equal(PowerPointSlideTransitionSpeed.Fast, slide.TransitionSpeed);
                    Assert.Equal(0.6, slide.TransitionDurationSeconds);
                    Assert.False(slide.TransitionAdvanceOnClick);
                    Assert.Equal(4.25, slide.TransitionAdvanceAfterSeconds);
                    Assert.Empty(presentation.ValidateDocument());
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.Single();
                    Transition transition = Assert.IsType<Transition>(slidePart.Slide.Transition);

                    Assert.Equal(TransitionSpeedValues.Fast, transition.Speed?.Value);
                    Assert.Equal("600", transition.Duration?.Value);
                    Assert.False(transition.AdvanceOnClick?.Value);
                    Assert.Equal("4250", transition.AdvanceAfterTime?.Value);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void TransitionMutation_PreservesTimingSettingsAcrossTransitionChanges() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    slide.Transition = PowerPointSlideTransition.Fade;
                    slide.TransitionSpeed = PowerPointSlideTransitionSpeed.Slow;
                    slide.TransitionDurationSeconds = 0.75;
                    slide.TransitionAdvanceAfterSeconds = 3.5;

                    slide.Transition = PowerPointSlideTransition.PushLeft;

                    presentation.Save();
                    Assert.Empty(presentation.ValidateDocument());
                }

                using (PowerPointDocument presentation = PowerPointDocument.Load(filePath)) {
                    PowerPointSlide slide = presentation.Slides.Single();
                    Assert.Equal(PowerPointSlideTransition.PushLeft, slide.Transition);
                    Assert.Equal(PowerPointSlideTransitionSpeed.Slow, slide.TransitionSpeed);
                    Assert.Equal(0.75, slide.TransitionDurationSeconds);
                    Assert.Equal(3.5, slide.TransitionAdvanceAfterSeconds);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void TransitionBackgroundAndNotesMutationsValidatePackage() {
            string filePath = CreateTempFilePath(".pptx");
            string imagesDirectory = AppendPathSegment(AppDomain.CurrentDomain.BaseDirectory, "Images");
            string imagePath = AppendPathSegment(imagesDirectory, "BackgroundImage.png");

            try {
                using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    slide.SetBackgroundImage(imagePath);
                    slide.Transition = PowerPointSlideTransition.Morph;
                    slide.Notes.Text = "Initial morph notes";

                    slide.Transition = PowerPointSlideTransition.Flash;
                    slide.BackgroundColor = OfficeColor.Parse("112233");
                    slide.Notes.Text = "Updated notes";

                    PowerPointSlide cleared = presentation.AddSlide();
                    cleared.Transition = PowerPointSlideTransition.Morph;
                    cleared.Transition = PowerPointSlideTransition.None;
                    cleared.SetBackgroundImage(imagePath);
                    cleared.ClearBackgroundImage();
                    cleared.Notes.Text = "Cleared transition notes";

                    presentation.Save();
                    Assert.Empty(presentation.ValidateDocument());
                }

                using (PowerPointDocument presentation = PowerPointDocument.Load(filePath)) {
                    Assert.Equal(PowerPointSlideTransition.Flash, presentation.Slides[0].Transition);
                    Assert.Equal(OfficeColor.Parse("112233"), presentation.Slides[0].BackgroundColor);
                    Assert.Equal("Updated notes", presentation.Slides[0].Notes.Text);
                    Assert.Equal(PowerPointSlideTransition.None, presentation.Slides[1].Transition);
                    Assert.Equal("Cleared transition notes", presentation.Slides[1].Notes.Text);
                    Assert.Empty(presentation.ValidateDocument());
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    OpenXmlValidator validator = new(FileFormatVersions.Microsoft365);
                    Assert.Empty(validator.Validate(document));

                    SlidePart[] slideParts = document.PresentationPart!.SlideParts.ToArray();
                    Assert.DoesNotContain("AlternateContent", slideParts[0].Slide.OuterXml, StringComparison.Ordinal);
                    Assert.Contains("flash", slideParts[0].Slide.OuterXml, StringComparison.OrdinalIgnoreCase);
                    Assert.DoesNotContain("AlternateContent", slideParts[1].Slide.OuterXml, StringComparison.Ordinal);
                    Assert.Null(slideParts[1].Slide.Transition);
                    Assert.Empty(slideParts[1].ImageParts);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void CanAddMultipleChartsWithUniqueAxisIds() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    for (int i = 0; i < 3; i++) {
                        slide.AddChart();
                    }

                    presentation.Save();
                    Assert.Empty(presentation.ValidateDocument());
                }

                using (PowerPointDocument presentation = PowerPointDocument.Load(filePath)) {
                    PowerPointSlide slide = presentation.Slides.Single();
                    Assert.Equal(3, slide.Charts.Count());
                    Assert.Empty(presentation.ValidateDocument());
                }

                using (PresentationDocument presentationDocument = PresentationDocument.Open(filePath, false)) {
                    PresentationPart? presentationPart = presentationDocument.PresentationPart;
                    Assert.NotNull(presentationPart);

                    HashSet<uint> axisIds = new();
                    foreach (ChartPart chartPart in presentationPart!.SlideParts.SelectMany(s => s.ChartParts)) {
                        Chart? chart = chartPart.ChartSpace?.GetFirstChild<Chart>();
                        Assert.NotNull(chart);

                        IEnumerable<uint> axisValues = (chart!.PlotArea?.Elements<OpenXmlCompositeElement>()
                            ?? Enumerable.Empty<OpenXmlCompositeElement>())
                            .Where(element => element is CategoryAxis || element is ValueAxis || element is SeriesAxis || element is DateAxis)
                            .SelectMany(element => element.Elements<AxisId>())
                            .Select(axis => axis.Val?.Value)
                            .OfType<uint>();

                        foreach (uint axisId in axisValues) {
                            Assert.True(axisIds.Add(axisId), $"Duplicate axis id {axisId} found.");
                        }
                    }

                    Assert.Equal(6, axisIds.Count);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void ChartAxisIdGenerator_DoesNotRewind_WhenAdvancedToOlderValue() {
            uint first = PowerPointChartAxisIdGenerator.GetNextId();
            uint second = PowerPointChartAxisIdGenerator.GetNextId();

            PowerPointChartAxisIdGenerator.AdvanceSeedToAtLeast(first);

            uint next = PowerPointChartAxisIdGenerator.GetNextId();
            Assert.True(next > second, $"Expected next axis id to stay above {second}, but got {next}.");
        }

        [Fact]
        public void ImportedMaximumAxisIdDoesNotExhaustLaterChartAuthoring() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                    presentation.AddSlide().AddChart(OfficeChartType.ColumnClustered,
                        new OfficeChartData(new[] { "A" }, new[] { new OfficeSeries("Values", new[] { 1d }) }));
                    ChartSpace chart = presentation.Slides[0].SlidePart.ChartParts.Single().ChartSpace!;
                    uint oldId = chart.Descendants<AxisId>().First().Val!.Value;
                    foreach (AxisId axis in chart.Descendants<AxisId>().Where(axis => axis.Val?.Value == oldId))
                        axis.Val = uint.MaxValue;
                    presentation.Save();
                }

                using PowerPointDocument imported = PowerPointDocument.Load(filePath);
                imported.AddSlide().AddChart(OfficeChartType.ColumnClustered,
                    new OfficeChartData(new[] { "A" }, new[] { new OfficeSeries("New", new[] { 2d }) }));
                Assert.All(imported.Slides.Last().SlidePart.ChartParts.Single().ChartSpace!.Descendants<AxisId>(),
                    axis => Assert.NotEqual(uint.MaxValue, axis.Val!.Value));
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }

        [Fact]
        public void ChartAddedToLoadedPresentationAllocatesAxisIdsAboveExistingCharts() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                uint existingId;
                using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                    presentation.AddSlide().AddChart(OfficeChartType.ColumnClustered,
                        new OfficeChartData(new[] { "A" }, new[] { new OfficeSeries("Values", new[] { 1d }) }));
                    ChartSpace chart = presentation.Slides[0].SlidePart.ChartParts.Single().ChartSpace!;
                    // Far ahead of the shared seed, so only reading the loaded deck can move past it.
                    existingId = PowerPointChartAxisIdGenerator.GetNextId() + 100000U;
                    uint[] oldIds = chart.Descendants<AxisId>().Select(axis => axis.Val!.Value).Distinct().ToArray();
                    foreach (AxisId axis in chart.Descendants<AxisId>())
                        axis.Val = existingId + (uint)Array.IndexOf(oldIds, axis.Val!.Value);
                    presentation.Save();
                }

                using PowerPointDocument loaded = PowerPointDocument.Load(filePath);
                loaded.AddSlide().AddChart(OfficeChartType.ColumnClustered,
                    new OfficeChartData(new[] { "A" }, new[] { new OfficeSeries("New", new[] { 2d }) }));
                Assert.All(loaded.Slides.Last().SlidePart.ChartParts.Single().ChartSpace!.Descendants<AxisId>(),
                    axis => Assert.True(axis.Val!.Value > existingId + 1U, $"Axis id {axis.Val!.Value} is not above the loaded deck's ids."));
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }

        [Fact]
        public void BackgroundColor_ReplacesExistingBackgroundImageFill() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");

            try {
                using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    slide.SetBackgroundImage(imagePath);
                    slide.BackgroundColor = OfficeColor.Parse("112233");

                    presentation.Save();
                    Assert.Empty(presentation.ValidateDocument());
                }

                using (PowerPointDocument presentation = PowerPointDocument.Load(filePath)) {
                    PowerPointSlide slide = presentation.Slides.Single();
                    Assert.Equal(OfficeColor.Parse("112233"), slide.BackgroundColor);
                    Assert.Empty(presentation.ValidateDocument());
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.First();
                    DocumentFormat.OpenXml.Presentation.BackgroundProperties properties =
                        slidePart.Slide.CommonSlideData!.Background!.BackgroundProperties!;

                    Assert.Null(properties.GetFirstChild<A.BlipFill>());
                    Assert.Equal("112233", properties.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void BackgroundGradient_ReplacesSolidFillAndValidatesPackage() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    slide.BackgroundColor = OfficeColor.Parse("112233");
                    slide.SetBackgroundGradient(OfficeColor.Parse("112233"), OfficeColor.Parse("445566"));

                    presentation.Save();
                    Assert.Empty(presentation.ValidateDocument());
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.First();
                    DocumentFormat.OpenXml.Presentation.BackgroundProperties properties =
                        slidePart.Slide.CommonSlideData!.Background!.BackgroundProperties!;

                    Assert.Null(properties.GetFirstChild<A.SolidFill>());
                    A.GradientFill gradient = Assert.IsType<A.GradientFill>(properties.GetFirstChild<A.GradientFill>());
                    A.GradientStop[] stops = gradient.GetFirstChild<A.GradientStopList>()!.Elements<A.GradientStop>().ToArray();
                    Assert.Equal(2, stops.Length);
                    Assert.Equal("112233", stops[0].GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);
                    Assert.Equal("445566", stops[1].GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);
                    Assert.Equal(8100000, gradient.GetFirstChild<A.LinearGradientFill>()?.Angle?.Value);

                    OpenXmlValidator validator = new(FileFormatVersions.Microsoft365);
                    Assert.Empty(validator.Validate(document));
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void BackgroundGradient_UsesProvidedAngleAndValidatesPackage() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    slide.SetBackgroundGradient(OfficeColor.Parse("112233"), OfficeColor.Parse("445566"), 45d);

                    presentation.Save();
                    Assert.Empty(presentation.ValidateDocument());
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.First();
                    DocumentFormat.OpenXml.Presentation.BackgroundProperties properties =
                        slidePart.Slide.CommonSlideData!.Background!.BackgroundProperties!;

                    A.GradientFill gradient = Assert.IsType<A.GradientFill>(properties.GetFirstChild<A.GradientFill>());
                    Assert.Equal(2700000, gradient.GetFirstChild<A.LinearGradientFill>()?.Angle?.Value);

                    OpenXmlValidator validator = new(FileFormatVersions.Microsoft365);
                    Assert.Empty(validator.Validate(document));
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void ClearingBackgroundColor_DoesNotRemoveBackgroundImage() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");

            try {
                using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    slide.SetBackgroundImage(imagePath);
                    slide.BackgroundColor = null;

                    presentation.Save();
                    Assert.Empty(presentation.ValidateDocument());
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.First();
                    DocumentFormat.OpenXml.Presentation.BackgroundProperties properties =
                        slidePart.Slide.CommonSlideData!.Background!.BackgroundProperties!;

                    Assert.NotNull(properties.GetFirstChild<A.BlipFill>());
                    Assert.Null(properties.GetFirstChild<A.SolidFill>());
                    Assert.Single(slidePart.ImageParts);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void ReplacingBackgroundImage_RemovesOrphanedImagePart() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string originalImage = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");
            string replacementImage = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "Kulek.jpg");

            try {
                using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    slide.SetBackgroundImage(originalImage);
                    slide.SetBackgroundImage(replacementImage);

                    presentation.Save();
                    Assert.Empty(presentation.ValidateDocument());
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.First();
                    ImagePart[] imageParts = slidePart.ImageParts.ToArray();

                    Assert.Single(imageParts);
                    Assert.Equal("image/jpeg", imageParts[0].ContentType);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void ClearingBackgroundImage_RemovesOrphanedImagePart() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");

            try {
                using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    slide.SetBackgroundImage(imagePath);
                    slide.ClearBackgroundImage();

                    presentation.Save();
                    Assert.Empty(presentation.ValidateDocument());
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.First();
                    DocumentFormat.OpenXml.Presentation.Background? background =
                        slidePart.Slide.CommonSlideData!.Background;

                    Assert.Empty(slidePart.ImageParts);
                    Assert.True(background == null ||
                        background.BackgroundProperties?.GetFirstChild<A.BlipFill>() == null);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void NotesText_ReadsAllParagraphsAndRuns() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    slide.Notes.Text = "seed";
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    DocumentFormat.OpenXml.Presentation.Shape notesShape = GetNotesTextShape(document);
                    notesShape.TextBody!.RemoveAllChildren<A.Paragraph>();
                    notesShape.TextBody.Append(
                        new A.Paragraph(
                            new A.Run(new A.Text("Alpha")),
                            new A.Run(new A.Text(" Beta"))),
                        new A.Paragraph(
                            new A.Run(new A.Text("Gamma")),
                            new A.Break(),
                            new A.Run(new A.Text("Delta"))));
                    document.PresentationPart!.SlideParts.First().NotesSlidePart!.NotesSlide!.Save();
                }

                using (PowerPointDocument presentation = PowerPointDocument.Load(filePath)) {
                    Assert.Equal($"Alpha Beta{Environment.NewLine}Gamma{Environment.NewLine}Delta",
                        presentation.Slides.Single().Notes.Text);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void NotesText_SetterReplacesExistingParagraphStructure() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    slide.Notes.Text = "Old";
                    presentation.Save();
                }

                using (PowerPointDocument presentation = PowerPointDocument.Load(filePath)) {
                    presentation.Slides.Single().Notes.Text = "First line\r\nSecond line";
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    DocumentFormat.OpenXml.Presentation.Shape notesShape = GetNotesTextShape(document);
                    A.Paragraph[] paragraphs = notesShape.TextBody!.Elements<A.Paragraph>().ToArray();

                    Assert.Equal(2, paragraphs.Length);
                    Assert.Equal("First line", ReadParagraphText(paragraphs[0]));
                    Assert.Equal("Second line", ReadParagraphText(paragraphs[1]));
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        private static DocumentFormat.OpenXml.Presentation.Shape GetNotesTextShape(PresentationDocument document) {
            return document.PresentationPart!.SlideParts.First().NotesSlidePart!.NotesSlide!.CommonSlideData!.ShapeTree!
                .Elements<DocumentFormat.OpenXml.Presentation.Shape>()
                .First(shape =>
                    shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<PlaceholderShape>()?.Type?.Value ==
                    PlaceholderValues.Body);
        }

        private static string ReadParagraphText(A.Paragraph paragraph) {
            StringBuilder builder = new();
            foreach (OpenXmlElement child in paragraph.ChildElements) {
                switch (child) {
                    case A.Run run:
                        builder.Append(run.Text?.Text ?? string.Empty);
                        break;
                    case A.Break:
                        builder.AppendLine();
                        break;
                    case A.Field field:
                        builder.Append(field.Text?.Text ?? string.Empty);
                        break;
                }
            }

            return builder.ToString();
        }

        private static string CreateTempFilePath(string extension) {
            string path = Path.GetTempFileName();
            File.Delete(path);
            return Path.ChangeExtension(path, extension);
        }

        private static string AppendPathSegment(string directoryPath, string childName) {
            string normalizedDirectory = Path.GetFullPath(directoryPath)
                .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            return normalizedDirectory + Path.DirectorySeparatorChar + childName;
        }
    }
}
