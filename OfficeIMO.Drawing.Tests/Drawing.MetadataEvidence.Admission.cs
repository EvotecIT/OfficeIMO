using System;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed partial class ImageMetadataEvidenceTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MetadataInputAdmissionRetainsTheConfiguredNativeCeiling(bool exifOnly) {
        byte[] bytes = exifOnly ? Exif(true, Field(OfficeExifDirectory.Image, 271, 2, Ascii("Camera")))
            : OfficeRasterImageEncoder.Encode(new OfficeRasterImage(1, 1, OfficeColor.Blue), OfficeImageExportFormat.Png);
        OfficeImageMetadataInspection Read(int maximum) {
            var options = new OfficeImageMetadataInspectionOptions { MaximumInputBytes = maximum };
            return exifOnly ? OfficeImageMetadata.InspectExifProfile(bytes, options) : OfficeImageMetadata.Inspect(bytes, options);
        }

        Assert.NotNull(Read(bytes.Length));
        OfficeInspectionResourceLimitException error = Assert.Throws<OfficeInspectionResourceLimitException>(() => Read(bytes.Length - 1));
        Assert.Equal("image-metadata", error.Owner);
        Assert.Equal(nameof(OfficeImageMetadataInspectionOptions.MaximumInputBytes), error.LimitName);
        Assert.Equal(bytes.Length - 1, error.Maximum);
    }
}
