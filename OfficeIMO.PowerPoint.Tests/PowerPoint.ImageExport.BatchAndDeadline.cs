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
        using var presentation = PowerPointPresentation.Create(new MemoryStream());
        presentation.SlideSize.SetSizePoints(80, 60);
        for (int i = 0; i < 3; i++) presentation.AddSlide();
        var options = new PowerPointPresentationImageExportOptions { SlideNumbers = new[] { 1, 3 } };
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
    }

    [Fact]
    public void VisualSnapshotEnforcesItsPublicTimeoutContract() {
        using var presentation = PowerPointPresentation.Create(new MemoryStream());
        var slide = presentation.AddSlide();
        Assert.Throws<OfficeImageExportTimeoutException>(() => slide.CreateVisualSnapshot(
            new PowerPointImageExportOptions { RenderTimeout = TimeSpan.FromTicks(1) }));
    }
}
