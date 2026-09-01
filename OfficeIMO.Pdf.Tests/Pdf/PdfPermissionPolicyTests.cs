using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfPermissionPolicyTests {
    [Fact]
    public void RestrictedUserPasswordBlocksTextUntilCallerExplicitlyIgnoresRestrictions() {
        byte[] pdf = CreateRestrictedPdf("open-one", "owner-one", "Restricted text");
        var enforced = new PdfLoadOptions { Password = "open-one" };
        var ignored = new PdfLoadOptions {
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
        var options = new PdfLoadOptions {
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
        var options = new PdfLoadOptions { Password = "owner-three" };

        PdfDocumentPreflight preflight = PdfInspector.Preflight(pdf, options);
        string text = PdfDocument.Load(pdf, options).Reader.Text();

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
        var enforced = new PdfLoadOptions { Password = "visual-open" };
        PdfReadDocument document = PdfReadDocument.Open(pdf, enforced);

        Assert.Throws<PdfPermissionDeniedException>(() => document.Pages[0].GetImages());
        Assert.Throws<PdfPermissionDeniedException>(() => document.Pages[0].GetImagePlacements());
        Assert.Throws<PdfPermissionDeniedException>(() => document.Pages[0].ToDrawing());
        Assert.Throws<PdfPermissionDeniedException>(() => PdfImageExtractor.ExtractImages(document));
        Assert.Throws<PdfPermissionDeniedException>(() => PdfImageExtractor.ExtractImagePlacements(document));
        Assert.Throws<PdfPermissionDeniedException>(() => PdfDocument.Load(pdf, enforced).Images.Find(new PdfPageRegion(1, 0D, 0D, 100D, 100D)));

        var ignored = new PdfLoadOptions {
            Password = "visual-open",
            PermissionPolicy = PdfPermissionPolicy.IgnoreRestrictions
        };
        PdfReadDocument authorized = PdfReadDocument.Open(pdf, ignored);
        Assert.NotEmpty(authorized.Pages[0].GetImages());
        Assert.NotEmpty(authorized.Pages[0].GetImagePlacements());
        Assert.NotEmpty(authorized.Pages[0].ToDrawing().Elements);
    }

    [Fact]
    public void RestrictedInteractionMapRequiresTextExtractionPermission() {
        byte[] pdf = CreateRestrictedPdf("interaction-open", "interaction-owner", "Restricted interaction text");
        var enforced = new PdfLoadOptions { Password = "interaction-open" };

        Assert.Throws<PdfPermissionDeniedException>(() =>
            PdfPageInteractionMap.Create(pdf, 1, readOptions: enforced));

        var ignored = new PdfLoadOptions {
            Password = "interaction-open",
            PermissionPolicy = PdfPermissionPolicy.IgnoreRestrictions
        };
        PdfPageInteractionMap map = PdfPageInteractionMap.Create(pdf, 1, readOptions: ignored);
        Assert.Contains("Restricted interaction text", string.Concat(map.TextRegions.Select(region => region.Text)), StringComparison.Ordinal);
    }

    [Fact]
    public void AccessibilityPermissionAllowsTextButNotTheFullLogicalObjectModel() {
        byte[] pdf = CreateEncryptedPdf(
            "accessible-open",
            "accessible-owner",
            PdfStandardPermissions.Accessibility,
            "Accessible text");
        var options = new PdfLoadOptions { Password = "accessible-open" };

        PdfDocumentPreflight preflight = PdfInspector.Preflight(pdf, options);
        string text = PdfTextExtractor.ExtractAllText(pdf, (PdfTextLayoutOptions?)null, options);
        PdfPermissionDeniedException exception = Assert.Throws<PdfPermissionDeniedException>(() =>
            PdfDocumentReadResult.Load(pdf, null, options));

        Assert.True(preflight.CanExtractText);
        Assert.False(preflight.CanReadLogicalObjects);
        Assert.Contains("Accessible text", text, StringComparison.Ordinal);
        Assert.Equal(PdfStandardPermissions.CopyContents, exception.Permission);
    }

    [Fact]
    public void RestrictedDirectReadModelRequiresCopyPermissionForLogicalContent() {
        var encryption = new PdfStandardEncryptionOptions("direct-open") {
            OwnerPassword = "direct-owner",
            AllowedPermissions = PdfStandardPermissions.None
        };
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                CreateOutlineFromHeadings = true,
                IncludeXmpMetadata = true
            }.SetSrgbOutputIntent().SetEncryption(encryption))
            .Meta(title: "Direct title", author: "OfficeIMO")
            .Bookmark("DirectAnchor")
            .H1("Direct logical heading")
            .Paragraph(paragraph => paragraph.LinkToBookmark("Jump", "DirectAnchor"))
            .TextField("Direct.Name", value: "Ada")
            .ToBytes();
        var enforced = new PdfLoadOptions { Password = "direct-open" };
        PdfDocumentPreflight restrictedPreflight = PdfInspector.Preflight(pdf, enforced);
        PdfReadDocument restricted = PdfReadDocument.Open(pdf, enforced);

        Assert.True(restrictedPreflight.CanRead);
        Assert.False(restrictedPreflight.CanReadLogicalObjects);
        Assert.Null(restrictedPreflight.DocumentInfo);
        Assert.Throws<PdfPermissionDeniedException>(() => restricted.Metadata);
        Assert.Throws<PdfPermissionDeniedException>(() => restricted.XmpMetadata);
        Assert.Throws<PdfPermissionDeniedException>(() => restricted.Outlines);
        Assert.Throws<PdfPermissionDeniedException>(() => restricted.NamedDestinations);
        Assert.Throws<PdfPermissionDeniedException>(() => restricted.FormFields);
        Assert.Throws<PdfPermissionDeniedException>(() => restricted.OutputIntents);
        Assert.Throws<PdfPermissionDeniedException>(() => restricted.RawStructure());
        Assert.Throws<PdfPermissionDeniedException>(() => restricted.Pages[0].GetLinkAnnotations());
        Assert.Throws<PdfPermissionDeniedException>(() => restricted.Pages[0].GetAnnotations());
        Assert.Throws<PdfPermissionDeniedException>(() => restricted.Pages[0].GetPageActions());
        PdfDocument restrictedDocument = PdfDocument.Load(pdf, enforced);
        Assert.Throws<PdfPermissionDeniedException>(() => restrictedDocument.Inspect());
        Assert.Throws<PdfPermissionDeniedException>(() => restrictedDocument.Analyze());
        Assert.Throws<PdfPermissionDeniedException>(() => restrictedDocument.Diagnostics());
        Assert.Throws<PdfPermissionDeniedException>(() => restrictedDocument.AnalyzeOptimization());
        Assert.Throws<PdfPermissionDeniedException>(() => restrictedDocument.Debug());

        var ignored = new PdfLoadOptions {
            Password = "direct-open",
            PermissionPolicy = PdfPermissionPolicy.IgnoreRestrictions
        };
        PdfDocumentPreflight authorizedPreflight = PdfInspector.Preflight(pdf, ignored);
        PdfReadDocument authorized = PdfReadDocument.Open(pdf, ignored);
        Assert.True(authorizedPreflight.CanReadLogicalObjects);
        Assert.Equal("Ada", Assert.Single(authorizedPreflight.DocumentInfo!.FormFields).Value);
        Assert.Equal("Direct title", authorized.Metadata.Title);
        Assert.Equal("Direct title", authorized.XmpMetadata?.Title);
        Assert.NotEmpty(authorized.Outlines);
        Assert.Contains(authorized.NamedDestinations, destination => destination.Name == "DirectAnchor");
        Assert.Equal("Ada", Assert.Single(authorized.FormFields).Value);
        Assert.NotEmpty(authorized.OutputIntents);
        Assert.NotEmpty(authorized.RawStructure().Objects);
        Assert.NotEmpty(authorized.Pages[0].GetLinkAnnotations());
        Assert.NotEmpty(authorized.Pages[0].GetAnnotations());
        Assert.Empty(authorized.Pages[0].GetPageActions());
        PdfDocument authorizedDocument = PdfDocument.Load(pdf, ignored);
        Assert.Equal("Ada", Assert.Single(authorizedDocument.Inspect().FormFields).Value);
        Assert.Equal("Ada", Assert.Single(authorizedDocument.Analyze().Info.FormFields).Value);
        Assert.Equal("Direct title", authorizedDocument.Diagnostics().Info?.Metadata.Title);
        Assert.Equal("Direct title", authorizedDocument.AnalyzeOptimization().Diagnostics.Info?.Metadata.Title);
        Assert.NotEmpty(authorizedDocument.Debug().Objects);
    }

    [Fact]
    public void MergeUsesPerSourcePasswordsAndReportsSecurityDecisions() {
        byte[] first = CreateRestrictedPdf("open-first", "owner-first", "First encrypted page");
        byte[] second = CreateRestrictedPdf("open-second", "owner-second", "Second encrypted page");
        var firstOptions = new PdfLoadOptions { Password = "owner-first" };
        var secondOptions = new PdfLoadOptions {
            Password = "open-second",
            PermissionPolicy = PdfPermissionPolicy.IgnoreRestrictions
        };

        PdfMergeResult result = PdfDocument.MergeWithReport(
            new PdfMergeOptions(),
            PdfDocument.Load(first, firstOptions),
            PdfDocument.Load(second, secondOptions));

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
        string mergedText = result.ToDocument().Reader.Text();
        Assert.Contains("First encrypted page", mergedText, StringComparison.Ordinal);
        Assert.Contains("Second encrypted page", mergedText, StringComparison.Ordinal);
    }

    [Fact]
    public void MergePreservesOriginalSecurityEvidenceAfterSourcePreprocessing() {
        byte[] encrypted = CreateRestrictedPdf("resize-open", "resize-owner", "Encrypted resized page");
        var sourceOptions = new PdfLoadOptions {
            Password = "resize-open",
            PermissionPolicy = PdfPermissionPolicy.IgnoreRestrictions
        };
        PdfDocument plain = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Plain page"));
        var mergeOptions = new PdfMergeOptions {
            ResizePages = new PdfPageResizeOptions(PageSizes.A4)
        };

        PdfMergeResult result = PdfDocument.MergeWithReport(
            mergeOptions,
            PdfDocument.Load(encrypted, sourceOptions),
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
        var restrictedOptions = new PdfLoadOptions { Password = "open-blocked" };
        PdfDocument plain = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Plain page"));

        PdfMutationBlockedException exception = Assert.Throws<PdfMutationBlockedException>(() =>
            PdfDocument.Merge(plain, PdfDocument.Load(restricted, restrictedOptions)));

        Assert.Contains("FullRewrite.Encryption", exception.Plan.BlockerCodes);

        byte[] allowed = CreateEncryptedPdf(
            "open-allowed",
            "owner-allowed",
            PdfStandardPermissions.CopyContents | PdfStandardPermissions.AssembleDocument,
            "Allowed merge");
        PdfDocument merged = PdfDocument.Merge(
            plain,
            PdfDocument.Load(allowed, new PdfLoadOptions { Password = "open-allowed" }));

        PdfMutationPlan allowedPlan = PdfMutationPlanner.Plan(
            allowed,
            PdfMutationOperation.MergeDocuments,
            new PdfLoadOptions { Password = "open-allowed" });

        Assert.Equal(2, PdfInspector.Inspect(merged.ToBytes()).PageCount);
        Assert.Contains("Allowed merge", merged.Reader.Text(), StringComparison.Ordinal);
        Assert.Contains(PdfMutationPermissionCheck.CopyContents, allowedPlan.PermissionChecks);
        Assert.Contains(PdfMutationPermissionCheck.AssembleDocument, allowedPlan.PermissionChecks);
        Assert.DoesNotContain(PdfMutationPermissionCheck.ModifyDocument, allowedPlan.PermissionChecks);
    }

    [Fact]
    public void TryMergeWithUsesExplicitReadOptionsForEncryptedTargetAcrossOverloads() {
        byte[] target = CreateRestrictedPdf("merge-open", "merge-owner", "Encrypted target");
        byte[] incoming = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Incoming page"))
            .ToBytes();
        var options = new PdfLoadOptions {
            Password = "merge-open",
            PermissionPolicy = PdfPermissionPolicy.IgnoreRestrictions
        };
        string path = Path.Combine(Path.GetTempPath(), $"officeimo-merge-{Guid.NewGuid():N}.pdf");

        try {
            File.WriteAllBytes(path, incoming);
            using var stream = new MemoryStream(incoming, writable: false);
            PdfOperationResult<PdfDocument>[] results = {
                PdfDocument.Load(target).TryMergeWith(PdfDocument.Load(incoming), options),
                PdfDocument.Load(target).TryMergeWith(incoming, options),
                PdfDocument.Load(target).TryMergeWith(path, options),
                PdfDocument.Load(target).TryMergeWith(stream, options)
            };

            Assert.All(results, result => {
                Assert.True(result.Succeeded, string.Join(Environment.NewLine, result.Diagnostics));
                Assert.Equal(2, result.RequireValue().Inspect().PageCount);
                string text = result.RequireValue().Reader.Text();
                Assert.Contains("Encrypted target", text, StringComparison.Ordinal);
                Assert.Contains("Incoming page", text, StringComparison.Ordinal);
            });
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void PageImportsUseIndependentSourceReadOptionsAcrossPlacements() {
        byte[] target = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Target page"))
            .ToBytes();
        byte[] source = CreateRestrictedThreePagePdf("import-open", "import-owner");
        var sourceReadOptions = new PdfLoadOptions {
            Password = "import-open",
            PermissionPolicy = PdfPermissionPolicy.IgnoreRestrictions
        };
        var importOptions = new PdfPageImportOptions {
            SourceReadOptions = sourceReadOptions
        };

        PdfDocument appended = PdfDocument.Load(target).Pages.Append(
            source,
            PdfPageSelection.From(2),
            importOptions);
        PdfDocument prepended = PdfDocument.Load(target).Pages.Prepend(source, importOptions);
        PdfDocument encryptedSourceDocument = PdfDocument.Load(source, sourceReadOptions);
        PdfOperationResult<PdfDocument> inserted = PdfDocument.Load(target).Pages.TryInsert(
            1,
            encryptedSourceDocument,
            PdfPageSelection.From(3),
            new PdfPageImportOptions());

        Assert.Equal(2, appended.Inspect().PageCount);
        Assert.Contains("Page two", appended.Reader.Text(), StringComparison.Ordinal);
        Assert.Equal(4, prepended.Inspect().PageCount);
        Assert.StartsWith("Page one", prepended.Reader.Text(), StringComparison.Ordinal);
        Assert.True(inserted.Succeeded, string.Join(Environment.NewLine, inserted.Diagnostics));
        Assert.Equal(2, inserted.RequireValue().Inspect().PageCount);
        Assert.StartsWith("Page three", inserted.RequireValue().Reader.Text(), StringComparison.Ordinal);
    }

    [Fact]
    public void AuthenticatedFluentPageMutationsPreserveStoredReadOptions() {
        byte[] source = CreateRestrictedThreePagePdf("pages-open", "pages-owner");
        var options = new PdfLoadOptions {
            Password = "pages-open",
            PermissionPolicy = PdfPermissionPolicy.IgnoreRestrictions
        };

        PdfDocument deleted = PdfDocument.Load(source, options).Pages.Delete(3);
        PdfDocument moved = PdfDocument.Load(source, options).Pages.Move(1, 3);
        PdfDocument rotated = PdfDocument.Load(source, options).Pages.Rotate(90, 1);
        PdfDocument boxed = PdfDocument.Load(source, options).Pages.SetCropBox(10, 10, 300, 500, 1);
        PdfDocument reordered = PdfDocument.Load(source, options).Pages.Reorder(3, 2, 1);
        PdfDocument duplicated = PdfDocument.Load(source, options).Pages.Duplicate(2);

        Assert.Equal(2, deleted.Inspect().PageCount);
        Assert.Equal(3, moved.Inspect().PageCount);
        Assert.StartsWith("Page three", moved.Reader.Text(), StringComparison.Ordinal);
        Assert.Equal(90, rotated.Inspect().Pages[0].RotationDegrees);
        Assert.Equal(290D, boxed.Inspect().Pages[0].CropBox!.Width, 3);
        Assert.StartsWith("Page three", reordered.Reader.Text(), StringComparison.Ordinal);
        Assert.Equal(4, duplicated.Inspect().PageCount);

        PdfOperationResult<PdfDocument> explicitOptions = PdfDocument.Load(source).Pages.TryRotate(180, PdfPageSelection.From(2), options);
        Assert.True(explicitOptions.Succeeded, string.Join(Environment.NewLine, explicitOptions.Diagnostics));
        Assert.Equal(180, explicitOptions.RequireValue().Inspect().Pages[1].RotationDegrees);

        PdfOperationResult<PdfDocument> explicitDuplicateOptions = PdfDocument.Load(source).Pages.TryDuplicate(PdfPageSelection.From(1), options);
        Assert.True(explicitDuplicateOptions.Succeeded, string.Join(Environment.NewLine, explicitDuplicateOptions.Diagnostics));
        Assert.Equal(4, explicitDuplicateOptions.RequireValue().Inspect().PageCount);

        PdfOperationResult<PdfDocument>[] selectorResults = {
            PdfDocument.Load(source).Pages.TryDelete(PdfPageSelector.Parse("last"), options),
            PdfDocument.Load(source).Pages.TryReorder(PdfPageSelector.Parse("last..1"), options),
            PdfDocument.Load(source).Pages.TryDuplicate(PdfPageSelector.Parse("1"), options),
            PdfDocument.Load(source).Pages.TryMove(1, PdfPageSelector.Parse("last"), options),
            PdfDocument.Load(source).Pages.TryRotate(270, PdfPageSelector.Parse("2"), options)
        };

        Assert.All(selectorResults, result => Assert.True(result.Succeeded, string.Join(Environment.NewLine, result.Diagnostics)));
        Assert.Equal(2, selectorResults[0].RequireValue().Inspect().PageCount);
        Assert.StartsWith("Page three", selectorResults[1].RequireValue().Reader.Text(), StringComparison.Ordinal);
        Assert.Equal(4, selectorResults[2].RequireValue().Inspect().PageCount);
        Assert.StartsWith("Page three", selectorResults[3].RequireValue().Reader.Text(), StringComparison.Ordinal);
        Assert.Equal(270, selectorResults[4].RequireValue().Inspect().Pages[1].RotationDegrees);
    }

    [Fact]
    public void AuthenticatedFluentFormMutationsPreserveStoredReadOptions() {
        byte[] source = CreateRestrictedFormPdf("forms-open", "forms-owner", "Before");
        var options = new PdfLoadOptions {
            Password = "forms-open",
            PermissionPolicy = PdfPermissionPolicy.IgnoreRestrictions
        };
        var ownerOptions = new PdfLoadOptions { Password = "forms-owner" };
        var values = new Dictionary<string, string> { ["Name"] = "After" };
        var data = new PdfFormDataSet(new[] { new PdfFormDataField("Name", new[] { "Imported" }) });

        PdfDocument filled = PdfDocument.Load(source, options).Forms.Fill(values);
        PdfDocument appended = PdfDocument.Load(source, ownerOptions).Forms.AppendRevision(values);
        PdfDocument flattened = PdfDocument.Load(source, options).Forms.Flatten();
        PdfDocument filledAndFlattened = PdfDocument.Load(source, options).Forms.FillAndFlatten(values);
        PdfDocument imported = PdfDocument.Load(source, options).Forms.ImportXfdf(data.ToXfdf());

        Assert.Equal("After", Assert.Single(filled.Inspect().FormFields).Value);
        Assert.Equal("After", Assert.Single(appended.Inspect().FormFields).Value);
        Assert.False(flattened.Inspect().HasForms);
        Assert.False(filledAndFlattened.Inspect().HasForms);
        Assert.Equal("Imported", Assert.Single(imported.Inspect().FormFields).Value);

        PdfOperationResult<PdfDocument> explicitOptions = PdfDocument.Load(source).Forms.TryFlatten(options);
        Assert.True(explicitOptions.Succeeded, string.Join(Environment.NewLine, explicitOptions.Diagnostics));
        Assert.False(explicitOptions.RequireValue().Inspect().HasForms);
    }

    [Fact]
    public void AuthenticatedSanitizationMetadataAndRewriteProofUseDocumentReadOptions() {
        byte[] source = CreateRestrictedPdf("owner-open", "owner-password", "Protected content");
        var options = new PdfLoadOptions {
            Password = "owner-open",
            PermissionPolicy = PdfPermissionPolicy.IgnoreRestrictions
        };

        PdfSanitizationResult sanitized = PdfDocument.Load(source, options).Sanitize();
        PdfDocument updated = PdfDocument.Load(source, options).UpdateMetadata(title: "Authenticated title");
        Assert.True(sanitized.IsSanitized);
        Assert.Equal("Authenticated title", updated.Reader.Metadata().Title);

        byte[] rewrittenBytes = CreateRestrictedPdf("rewrite-open", "rewrite-owner", "Protected content");
        var rewrittenOptions = new PdfLoadOptions {
            Password = "rewrite-open",
            PermissionPolicy = PdfPermissionPolicy.IgnoreRestrictions
        };
        PdfRewritePreservationReport report = PdfDocument.Load(source, options).AssessRewritePreservation(
            PdfDocument.Load(rewrittenBytes, rewrittenOptions));

        Assert.True(report.Original.Security.HasEncryption);
        Assert.True(report.Rewritten.Security.HasEncryption);

        var enforcedSourceOptions = new PdfLoadOptions { Password = "owner-open" };
        var enforcedRewrittenOptions = new PdfLoadOptions { Password = "rewrite-open" };
        Assert.Throws<PdfPermissionDeniedException>(() =>
            PdfDocument.Load(source, enforcedSourceOptions).AssessRewritePreservation(
                PdfDocument.Load(rewrittenBytes, rewrittenOptions)));
        Assert.Throws<PdfPermissionDeniedException>(() =>
            PdfDocument.Load(source, options).AssessRewritePreservation(
                PdfDocument.Load(rewrittenBytes, enforcedRewrittenOptions)));
    }

    private static byte[] CreateRestrictedPdf(string userPassword, string ownerPassword, string text) =>
        CreateEncryptedPdf(userPassword, ownerPassword, PdfStandardPermissions.None, text);

    private static byte[] CreateRestrictedThreePagePdf(string userPassword, string ownerPassword) {
        var encryption = new PdfStandardEncryptionOptions(userPassword) {
            OwnerPassword = ownerPassword,
            AllowedPermissions = PdfStandardPermissions.None
        };
        return PdfDocument.Create(new PdfOptions().SetEncryption(encryption))
            .Paragraph(paragraph => paragraph.Text("Page one"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Page two"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Page three"))
            .ToBytes();
    }

    private static byte[] CreateRestrictedFormPdf(string userPassword, string ownerPassword, string value) {
        var encryption = new PdfStandardEncryptionOptions(userPassword) {
            OwnerPassword = ownerPassword,
            AllowedPermissions = PdfStandardPermissions.None
        };
        return PdfDocument.Create(new PdfOptions().SetEncryption(encryption))
            .TextField("Name", width: 180, height: 24, value: value)
            .ToBytes();
    }

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
