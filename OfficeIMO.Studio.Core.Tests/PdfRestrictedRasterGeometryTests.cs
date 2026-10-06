using System.Text;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Tests;

public sealed class PdfRestrictedRasterGeometryTests {
    [Theory]
    [InlineData(0)]
    [InlineData(90)]
    public async Task RestrictedRasterUsesPhysicalPageGeometryForDeviceScaleAndBudget(int rotation) {
        string path = Path.Combine(Path.GetTempPath(), $"officeimo-restricted-geometry-{Guid.NewGuid():N}.pdf");
        byte[] source = Encoding.ASCII.GetBytes($"""
            %PDF-1.7
            1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj
            2 0 obj << /Type /Pages /Count 1 /Kids [3 0 R] >> endobj
            3 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /UserUnit 2 /Rotate {rotation} /Resources << >> >> endobj
            trailer << /Root 1 0 R /Size 4 >>
            %%EOF
            """);
        byte[] encrypted = PdfDocument.Load(source).Security.Encrypt(new PdfStandardEncryptionOptions("reader") {
            OwnerPassword = "owner", AllowedPermissions = PdfStandardPermissions.None
        }).Pdf;
        await File.WriteAllBytesAsync(path, encrypted);
        try {
            using var workspace = await PdfWorkspace.OpenAsync(path, default, password: "reader");
            PdfDocumentSession session = PdfDocumentSession.FromWorkspace(workspace);
            PdfPageScene scene = await session.LoadPageSceneAsync(1, default);
            Assert.True(scene.RequiresRasterFallback);
            Assert.Null(scene.Interactions);
            Assert.Empty(scene.Drawing.Elements);
            double width = rotation == 90 ? 1584D : 1224D;
            double height = rotation == 90 ? 1224D : 1584D;
            // A 100% page box has half as many display units as the physical drawing at UserUnit 2.
            double scale = PdfRasterScale.Compose(0.5D, 3D, scene.Drawing.Width, scene.Drawing.Height,
                StudioPdfSecurityPolicy.MaximumRasterPixels);
            Assert.Equal(width, scene.Drawing.Width);
            Assert.Equal(height, scene.Drawing.Height);
            Assert.Equal(1.5D, scale);
            PdfRenderedPage rendered = await session.RenderPageAsync(1, scale, default);
            Assert.Equal((int)(width * scale), rendered.PixelWidth);
            Assert.Equal((int)(height * scale), rendered.PixelHeight);
            Assert.True((long)rendered.PixelWidth * rendered.PixelHeight <= StudioPdfSecurityPolicy.MaximumRasterPixels);
        } finally { File.Delete(path); }
    }
}
