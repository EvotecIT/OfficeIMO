using System;
using System.Collections.Generic;
using System.IO;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

public partial class PowerPointImageExportTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SelectedSlideExportsRetainOriginalIdentityAndBatchPosition(bool streaming) {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(80),OfficeIMO.OfficeLength.Points(60));
        for (int i = 0; i < 3; i++) presentation.AddSlide();
        var options = new PowerPointImageExportOptions { SlideNumbers = new[] { 1, 3 } };
        var results = new List<OfficeImageExportResult>();
        if (streaming) presentation.ExportImages(OfficeImageExportFormat.Svg, results.Add, options);
        else results.AddRange(presentation.ExportImages(OfficeImageExportFormat.Svg, options));
        Assert.Equal(2, results.Count);
        for (int index = 0; index < results.Count; index++) {
            Assert.Equal(index, results[index].SequenceIndex);
            Assert.Equal(2, results[index].SequenceCount);
            Assert.Equal("Slide " + (index == 0 ? 1 : 3), results[index].Name);
            Assert.Equal("PowerPoint slide " + (index == 0 ? 1 : 3), results[index].Source);
            Assert.NotEmpty(results[index].Bytes);
        }

        options.MaximumOutputCount = 1;
        results.Clear();
        var failure = Assert.Throws<OfficeImageExportBatchLimitException>(() => {
            if (streaming) presentation.ExportImages(OfficeImageExportFormat.Svg, results.Add, options);
            else results.AddRange(presentation.ExportImages(OfficeImageExportFormat.Svg, options));
        });
        Assert.Equal(nameof(options.MaximumOutputCount), failure.LimitName);
        Assert.Empty(results);

        options.MaximumOutputCount = 2;
        options.SlideNumbers = new[] { 3, 1, 3 };
        Assert.Equal(2, presentation.ExportImages(OfficeImageExportFormat.Svg, options).Count);
        presentation.Slides[1].Hidden = true;
        options.SlideNumbers = null;
        Assert.Equal(2, presentation.ExportImages(OfficeImageExportFormat.Svg, options).Count);
        options.IncludeHiddenSlides = true;
        results.Clear();
        Assert.Throws<OfficeImageExportBatchLimitException>(() => presentation.ExportImages(
            OfficeImageExportFormat.Svg, results.Add, options));
        Assert.Empty(results);
    }

    [Fact]
    public void VisualSnapshotEnforcesItsPublicTimeoutContract() {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        var slide = presentation.AddSlide();
        Assert.Throws<OfficeImageExportTimeoutException>(() => slide.CreateVisualSnapshot(
            new PowerPointImageExportOptions { RenderTimeout = TimeSpan.FromTicks(1) }));
    }
}
