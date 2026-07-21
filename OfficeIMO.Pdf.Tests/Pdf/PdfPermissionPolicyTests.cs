using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfPermissionPolicyTests {
    [Fact]
    public void RestrictedUserPasswordBlocksTextUntilCallerExplicitlyIgnoresRestrictions() {
        byte[] pdf = CreateRestrictedPdf("open-one", "owner-one", "Restricted text");
        var enforced = new PdfReadOptions { Password = "open-one" };
        var ignored = new PdfReadOptions {
            Password = "open-one",
            PermissionPolicy = PdfPermissionPolicy.IgnoreRestrictions
        };

        PdfDocumentPreflight enforcedPreflight = PdfInspector.Preflight(pdf, enforced);
        PdfPermissionDeniedException exception = Assert.Throws<PdfPermissionDeniedException>(() =>
            PdfTextExtractor.ExtractAllText(pdf, (PdfTextLayoutOptions?)null, enforced));
        PdfDocumentPreflight ignoredPreflight = PdfInspector.Preflight(pdf, ignored);
        string text = PdfTextExtractor.ExtractAllText(pdf, (PdfTextLayoutOptions?)null, ignored);

        Assert.True(enforcedPreflight.CanRead, string.Join(Environment.NewLine, enforcedPreflight.Diagnostics));
        Assert.False(enforcedPreflight.CanExtractText);
        Assert.Equal(PdfStandardPermissions.CopyContents, exception.Permission);
        Assert.Equal(PdfPasswordAuthenticationRole.User, exception.AuthenticationRole);
        Assert.True(ignoredPreflight.CanExtractText);
        Assert.True(ignoredPreflight.PermissionRestrictionsIgnored);
        Assert.True(ignoredPreflight.CanManipulatePages);
        Assert.Contains("Restricted text", text, StringComparison.Ordinal);
    }

    [Fact]
    public void IgnoreRestrictionsStillRequiresAValidDecryptionPassword() {
        byte[] pdf = CreateRestrictedPdf("open-two", "owner-two", "No password bypass");
        var options = new PdfReadOptions {
            Password = "wrong",
            PermissionPolicy = PdfPermissionPolicy.IgnoreRestrictions
        };

        Assert.Throws<PdfInvalidPasswordException>(() => PdfReadDocument.Open(pdf, options));
        Assert.Throws<PdfInvalidPasswordException>(() =>
            PdfTextExtractor.ExtractAllText(pdf, (PdfTextLayoutOptions?)null, options));
    }

    [Fact]
    public void OwnerPasswordDoesNotNeedPermissionOverride() {
        byte[] pdf = CreateRestrictedPdf("open-three", "owner-three", "Owner authorized text");
        var options = new PdfReadOptions { Password = "owner-three" };

        PdfDocumentPreflight preflight = PdfInspector.Preflight(pdf, options);
        string text = PdfDocument.Open(pdf, options).Read.Text();

        Assert.Equal(PdfPasswordAuthenticationRole.Owner, preflight.Probe.Security.PasswordAuthenticationRole);
        Assert.True(preflight.CanExtractText);
        Assert.False(preflight.PermissionRestrictionsIgnored);
        Assert.Contains("Owner authorized text", text, StringComparison.Ordinal);
    }

    [Fact]
    public void RestrictedPageLevelVisualExtractionRequiresCopyPermission() {
        var encryption = new PdfStandardEncryptionOptions("visual-open") {
            OwnerPassword = "visual-owner",
            AllowedPermissions = PdfStandardPermissions.None
        };
        byte[] pdf = PdfDocument.Create(new PdfOptions().SetEncryption(encryption))
            .Canvas(canvas => canvas.Image(PdfPngTestImages.CreateRgbPng(30, 90, 180), 20D, 20D, 40D, 40D))
            .ToBytes();
        var enforced = new PdfReadOptions { Password = "visual-open" };
        PdfReadDocument document = PdfReadDocument.Open(pdf, enforced);

        Assert.Throws<PdfPermissionDeniedException>(() => document.Pages[0].GetImages());
        Assert.Throws<PdfPermissionDeniedException>(() => document.Pages[0].GetImagePlacements());
        Assert.Throws<PdfPermissionDeniedException>(() => document.Pages[0].ToDrawing());
        Assert.Throws<PdfPermissionDeniedException>(() => PdfImageExtractor.ExtractImages(document));
        Assert.Throws<PdfPermissionDeniedException>(() => PdfImageExtractor.ExtractImagePlacements(document));

        var ignored = new PdfReadOptions {
            Password = "visual-open",
            PermissionPolicy = PdfPermissionPolicy.IgnoreRestrictions
        };
        PdfReadDocument authorized = PdfReadDocument.Open(pdf, ignored);
        Assert.NotEmpty(authorized.Pages[0].GetImages());
        Assert.NotEmpty(authorized.Pages[0].GetImagePlacements());
        Assert.NotEmpty(authorized.Pages[0].ToDrawing().Elements);
    }

    [Fact]
    public void AccessibilityPermissionAllowsTextButNotTheFullLogicalObjectModel() {
        byte[] pdf = CreateEncryptedPdf(
            "accessible-open",
            "accessible-owner",
            PdfStandardPermissions.Accessibility,
            "Accessible text");
        var options = new PdfReadOptions { Password = "accessible-open" };

        PdfDocumentPreflight preflight = PdfInspector.Preflight(pdf, options);
        string text = PdfTextExtractor.ExtractAllText(pdf, (PdfTextLayoutOptions?)null, options);
        PdfPermissionDeniedException exception = Assert.Throws<PdfPermissionDeniedException>(() =>
            PdfLogicalDocument.Load(pdf, null, options));

        Assert.True(preflight.CanExtractText);
        Assert.False(preflight.CanReadLogicalObjects);
        Assert.Contains("Accessible text", text, StringComparison.Ordinal);
        Assert.Equal(PdfStandardPermissions.CopyContents, exception.Permission);
    }

    [Fact]
    public void MergeUsesPerSourcePasswordsAndReportsSecurityDecisions() {
        byte[] first = CreateRestrictedPdf("open-first", "owner-first", "First encrypted page");
        byte[] second = CreateRestrictedPdf("open-second", "owner-second", "Second encrypted page");
        var firstOptions = new PdfReadOptions { Password = "owner-first" };
        var secondOptions = new PdfReadOptions {
            Password = "open-second",
            PermissionPolicy = PdfPermissionPolicy.IgnoreRestrictions
        };

        PdfMergeResult result = PdfDocument.MergeWithReport(
            new PdfMergeOptions(),
            PdfDocument.Open(first, firstOptions),
            PdfDocument.Open(second, secondOptions));

        Assert.Equal(2, result.Report.OutputPageCount);
        Assert.False(result.Report.OutputHasEncryption);
        Assert.False(result.Report.OutputHasSignatures);
        Assert.Equal(PdfPasswordAuthenticationRole.Owner, result.Report.Sources[0].PasswordAuthenticationRole);
        Assert.False(result.Report.Sources[0].PermissionRestrictionsIgnored);
        Assert.Equal(PdfPasswordAuthenticationRole.User, result.Report.Sources[1].PasswordAuthenticationRole);
        Assert.True(result.Report.Sources[1].PermissionRestrictionsIgnored);
        Assert.Equal(PdfStandardPermissions.None, result.Report.Sources[1].Security.AllowedStandardPermissions);
        PdfMergeDecision security = Assert.Single(result.Report.Decisions, decision => decision.Structure == "Security");
        Assert.Contains("unencrypted", security.Action, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("explicitly ignored", security.Action, StringComparison.OrdinalIgnoreCase);
        string mergedText = result.ToDocument().Read.Text();
        Assert.Contains("First encrypted page", mergedText, StringComparison.Ordinal);
        Assert.Contains("Second encrypted page", mergedText, StringComparison.Ordinal);
    }

    [Fact]
    public void MergePreservesOriginalSecurityEvidenceAfterSourcePreprocessing() {
        byte[] encrypted = CreateRestrictedPdf("resize-open", "resize-owner", "Encrypted resized page");
        var sourceOptions = new PdfReadOptions {
            Password = "resize-open",
            PermissionPolicy = PdfPermissionPolicy.IgnoreRestrictions
        };
        PdfDocument plain = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Plain page"));
        var mergeOptions = new PdfMergeOptions {
            ResizePages = new PdfPageResizeOptions(PageSizes.A4)
        };

        PdfMergeResult result = PdfDocument.MergeWithReport(
            mergeOptions,
            PdfDocument.Open(encrypted, sourceOptions),
            plain);

        PdfMergeSourceInventory inventory = result.Report.Sources[0];
        Assert.True(inventory.HasEncryption);
        Assert.Equal(PdfPasswordAuthenticationRole.User, inventory.PasswordAuthenticationRole);
        Assert.True(inventory.PermissionRestrictionsIgnored);
        Assert.Equal(PdfStandardPermissions.None, inventory.Security.AllowedStandardPermissions);
        PdfMergeDecision security = Assert.Single(result.Report.Decisions, decision => decision.Structure == "Security");
        Assert.Contains("unencrypted", security.Action, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("explicitly ignored", security.Action, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void RestrictedUserMergeIsBlockedUnlessCopyAndAssemblyAreAllowed() {
        byte[] restricted = CreateRestrictedPdf("open-blocked", "owner-blocked", "Blocked merge");
        var restrictedOptions = new PdfReadOptions { Password = "open-blocked" };
        PdfDocument plain = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Plain page"));

        PdfMutationBlockedException exception = Assert.Throws<PdfMutationBlockedException>(() =>
            PdfDocument.Merge(plain, PdfDocument.Open(restricted, restrictedOptions)));

        Assert.Contains("FullRewrite.Encryption", exception.Plan.BlockerCodes);

        byte[] allowed = CreateEncryptedPdf(
            "open-allowed",
            "owner-allowed",
            PdfStandardPermissions.CopyContents | PdfStandardPermissions.AssembleDocument,
            "Allowed merge");
        PdfDocument merged = PdfDocument.Merge(
            plain,
            PdfDocument.Open(allowed, new PdfReadOptions { Password = "open-allowed" }));

        PdfMutationPlan allowedPlan = PdfMutationPlanner.Plan(
            allowed,
            PdfMutationOperation.MergeDocuments,
            new PdfReadOptions { Password = "open-allowed" });

        Assert.Equal(2, PdfInspector.Inspect(merged.ToBytes()).PageCount);
        Assert.Contains("Allowed merge", merged.Read.Text(), StringComparison.Ordinal);
        Assert.Contains(PdfMutationPermissionCheck.CopyContents, allowedPlan.PermissionChecks);
        Assert.Contains(PdfMutationPermissionCheck.AssembleDocument, allowedPlan.PermissionChecks);
        Assert.DoesNotContain(PdfMutationPermissionCheck.ModifyDocument, allowedPlan.PermissionChecks);
    }

    private static byte[] CreateRestrictedPdf(string userPassword, string ownerPassword, string text) =>
        CreateEncryptedPdf(userPassword, ownerPassword, PdfStandardPermissions.None, text);

    private static byte[] CreateEncryptedPdf(
        string userPassword,
        string ownerPassword,
        PdfStandardPermissions permissions,
        string text) {
        var encryption = new PdfStandardEncryptionOptions(userPassword) {
            OwnerPassword = ownerPassword,
            AllowedPermissions = permissions
        };
        return PdfDocument.Create(new PdfOptions().SetEncryption(encryption))
            .Paragraph(paragraph => paragraph.Text(text))
            .ToBytes();
    }
}
