namespace OfficeIMO.Pdf.Testing.Tests;

public sealed class CommittedBaselineTests {
    /// <summary>
    /// The baseline in __pdfsnapshots__ was approved on macOS arm64. Running it on Linux and Windows checks that the
    /// default tolerances absorb raster differences between platforms while the structure file stays exact.
    /// </summary>
    [Fact]
    public void BaselineApprovedOnAnotherPlatformStillMatches() {
        PdfSnapshot.Verify(SampleReport.Create("1,250.00", badge: OfficeColor.FromRgb(0, 128, 0), pages: 2), "committed-report");
    }
}
