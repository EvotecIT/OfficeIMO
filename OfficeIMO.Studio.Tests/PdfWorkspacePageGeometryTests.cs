using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Tests;

public sealed partial class PdfWorkspaceTests {
    [Fact]
    public async Task NormalizeRotationRetainsAppearanceSizeAndCanBeUndone() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-normalize-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string source = Path.Combine(root, "source.pdf"), output = Path.Combine(root, "output.pdf");
            PdfDocument.Create(compose => compose.Page(page => page.Size(global::OfficeIMO.OfficeLength.Points(300), global::OfficeIMO.OfficeLength.Points(400))
                .Content(content => content.Text("Rotation normalization")))).Pages.RotateBy(90).Save(source);
            using var workspace = await PdfWorkspace.OpenAsync(source, CancellationToken.None,
                new PdfWorkspaceRecoveryStore(Path.Combine(root, "recovery")));
            Assert.True(workspace.CanNormalizePageRotation);
            await workspace.NormalizePageRotationAsync([1], CancellationToken.None);
            Assert.Equal(PdfWorkspaceOperationKind.NormalizeRotation, Assert.Single(workspace.Journal).Kind);
            Assert.Equal(0, workspace.Pages[0].RotationDegrees);
            await workspace.SaveAsync(output, CancellationToken.None);
            Assert.Equal(400D, PdfDocument.Load(output).GetPageLayouts()[0].VisualWidth);
            await workspace.UndoAsync(CancellationToken.None);
            Assert.Equal(90, workspace.Pages[0].RotationDegrees);
            await workspace.RedoAsync(CancellationToken.None);
            Assert.Equal(0, workspace.Pages[0].RotationDegrees);
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Fact]
    public async Task TaggedCropRetainsTheSourceAndResetRevealsTheFullPage() {
        string source = Path.Combine(AppContext.BaseDirectory, "Fixtures", "openpreserve-pdfa1b-text.pdf");
        string root = Path.Combine(Path.GetTempPath(), "officeimo-crop-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            byte[] original = File.ReadAllBytes(source);
            var recovery = new PdfWorkspaceRecoveryStore(Path.Combine(root, "recovery"));
            using PdfWorkspace workspace = await PdfWorkspace.OpenAsync(source, CancellationToken.None, recovery);
            Assert.True(workspace.CanCropPages);
            double width = workspace.Pages[0].Geometry.EffectiveBox!.Width;
            await workspace.CropByMarginAsync([1], 20, CancellationToken.None);
            Assert.Equal(width - 40, workspace.Pages[0].Geometry.EffectiveBox!.Width, 6);
            Assert.Equal(PdfWorkspaceOperationKind.CropBox, Assert.Single(workspace.Journal).Kind);
            string croppedPath = Path.Combine(root, "cropped.pdf");
            await workspace.SaveAsync(croppedPath, CancellationToken.None);
            byte[] cropped = File.ReadAllBytes(croppedPath);
            Assert.True(cropped.AsSpan(0, original.Length).SequenceEqual(original));
            Assert.Equal(PdfDocument.Load(new System.IO.MemoryStream(original)).Inspect().Security.RevisionCount + 1,
                PdfDocument.Load(new System.IO.MemoryStream(cropped)).Inspect().Security.RevisionCount);
            await workspace.ResetCropAsync([1], CancellationToken.None);
            Assert.Equal(width, workspace.Pages[0].Geometry.EffectiveBox!.Width, 6);
            await workspace.UndoAsync(CancellationToken.None);
            Assert.Equal(width - 40, workspace.Pages[0].Geometry.EffectiveBox!.Width, 6);
            await workspace.RedoAsync(CancellationToken.None);
            Assert.Equal(width, workspace.Pages[0].Geometry.EffectiveBox!.Width, 6);
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Fact]
    public async Task MultiplePageMarginCropIsOneUndoableOperation() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-crop-batch-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string source = Path.Combine(root, "source.pdf");
            PdfDocument.Create(compose => {
                for (int i = 0; i < 3; i++) compose.Page(page => page.Size(global::OfficeIMO.OfficeLength.Points(300), global::OfficeIMO.OfficeLength.Points(400)));
            }).Save(source);
            using PdfWorkspace workspace = await PdfWorkspace.OpenAsync(source, CancellationToken.None,
                new PdfWorkspaceRecoveryStore(Path.Combine(root, "recovery")));
            await workspace.CropByMarginAsync([1, 3], 20, CancellationToken.None);
            Assert.Single(workspace.Journal);
            Assert.Equal(new[] { 260D, 300D, 260D }, workspace.Pages.Select(page => page.Geometry.EffectiveBox!.Width));
            await workspace.UndoAsync(CancellationToken.None);
            Assert.All(workspace.Pages, page => Assert.Equal(300D, page.Geometry.EffectiveBox!.Width));
            Assert.False(workspace.IsDirty);
        } finally { Directory.Delete(root, recursive: true); }
    }
}
