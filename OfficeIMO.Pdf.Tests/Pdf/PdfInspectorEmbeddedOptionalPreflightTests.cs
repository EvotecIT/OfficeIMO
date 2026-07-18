using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfInspectorTests {
    [Fact]
    public void Preflight_AllowsSimpleEmbeddedFilePdfReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildEmbeddedFilePdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasEmbeddedFiles);
        Assert.True(report.Probe.HasCatalogNameTrees);
        Assert.False(report.Probe.HasOptionalContent);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasEmbeddedFiles);
        Assert.True(report.DocumentInfo.HasCatalogNameTrees);
        Assert.False(report.DocumentInfo.HasOptionalContent);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.RewriteBlockers);
        Assert.DoesNotContain("PDF embedded files are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
    }

    [Fact]
    public void Preflight_AllowsSimpleAssociatedFilePdfReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildAssociatedFilePdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasEmbeddedFiles);
        Assert.False(report.Probe.HasCatalogNameTrees);
        Assert.False(report.Probe.HasOptionalContent);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasEmbeddedFiles);
        Assert.False(report.DocumentInfo.HasCatalogNameTrees);
        Assert.False(report.DocumentInfo.HasOptionalContent);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.RewriteBlockers);
        Assert.DoesNotContain("PDF embedded files are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
    }

    [Fact]
    public void Preflight_AllowsCombinedDestinationAndEmbeddedFileNameTreesReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildCombinedDestinationAndEmbeddedFileNameTreePdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasNamedDestinations);
        Assert.True(report.Probe.HasEmbeddedFiles);
        Assert.True(report.Probe.HasCatalogNameTrees);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasNamedDestinations);
        Assert.True(report.DocumentInfo.HasEmbeddedFiles);
        Assert.True(report.DocumentInfo.HasCatalogNameTrees);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.RewriteBlockers);
        Assert.DoesNotContain("PDF named destinations are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        Assert.DoesNotContain("PDF embedded files are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
    }

    [Fact]
    public void Preflight_AllowsUnsupportedCatalogNameTreePdfReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildUnsupportedCatalogNameTreePdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasCatalogNameTrees);
        Assert.False(report.Probe.HasNamedDestinations);
        Assert.False(report.Probe.HasEmbeddedFiles);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasCatalogNameTrees);
        Assert.False(report.DocumentInfo.HasNamedDestinations);
        Assert.False(report.DocumentInfo.HasEmbeddedFiles);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("PDF catalog name trees are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.CatalogNameTrees, "PDF catalog name trees are not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Preflight_AllowsComplexEmbeddedFilePdfReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildComplexEmbeddedFilePdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasEmbeddedFiles);
        Assert.True(report.Probe.HasCatalogNameTrees);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasEmbeddedFiles);
        Assert.True(report.DocumentInfo.HasCatalogNameTrees);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("PDF embedded files are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.EmbeddedFiles, "PDF embedded files are not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Preflight_AllowsComplexAssociatedFilePdfReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildComplexAssociatedFilePdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasEmbeddedFiles);
        Assert.False(report.Probe.HasCatalogNameTrees);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasEmbeddedFiles);
        Assert.False(report.DocumentInfo.HasCatalogNameTrees);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("PDF embedded files are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.EmbeddedFiles, "PDF embedded files are not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Inspect_ReadsEmbeddedFileAttachmentMetadataWithoutPayloads() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildEmbeddedFilePdf());

        Assert.True(report.CanRead);
        Assert.True(report.Probe.HasEmbeddedFiles);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasEmbeddedFiles);
        Assert.True(report.DocumentInfo.HasAttachments);
        Assert.Equal(1, report.DocumentInfo.AttachmentCount);
        Assert.Equal(new[] { "note.txt" }, report.DocumentInfo.AttachmentNames);
        Assert.Equal(new[] { "note.txt" }, report.DocumentInfo.AttachmentFileNames);
        Assert.Equal(new[] { "Names/EmbeddedFiles" }, report.DocumentInfo.AttachmentSources);

        PdfAttachmentInfo attachment = Assert.Single(report.DocumentInfo.Attachments);
        Assert.Equal("note.txt", attachment.Name);
        Assert.Equal("note.txt", attachment.FileName);
        Assert.Null(attachment.UnicodeFileName);
        Assert.Null(attachment.Description);
        Assert.Null(attachment.MimeType);
        Assert.Equal(PdfAssociatedFileRelationship.Unspecified, attachment.Relationship);
        Assert.Equal(4, attachment.SizeBytes);
        Assert.Equal(5, attachment.FileSpecObjectNumber);
        Assert.Equal(6, attachment.EmbeddedFileObjectNumber);
        Assert.False(attachment.IsAssociatedFile);
        Assert.Same(attachment, Assert.Single(report.DocumentInfo.GetAttachmentsByName("note.txt")));
        Assert.Same(attachment, Assert.Single(report.DocumentInfo.GetAttachmentsByFileName("note.txt")));
        Assert.Same(attachment, Assert.Single(report.DocumentInfo.GetAttachmentsBySource("Names/EmbeddedFiles")));
        Assert.Same(attachment, Assert.Single(report.DocumentInfo.GetAttachmentsByRelationship(PdfAssociatedFileRelationship.Unspecified)));
        Assert.Empty(report.DocumentInfo.GetAttachmentsBySource("AF"));
    }

    [Fact]
    public void Inspect_ReadsAssociatedFileAttachmentMetadataWithoutNameTree() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildAssociatedFilePdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasEmbeddedFiles);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasEmbeddedFiles);
        Assert.True(report.DocumentInfo.HasAttachments);
        Assert.Equal(1, report.DocumentInfo.AttachmentCount);
        Assert.Equal(new[] { "data.xml" }, report.DocumentInfo.AttachmentNames);
        Assert.Equal(new[] { "AF" }, report.DocumentInfo.AttachmentSources);

        PdfAttachmentInfo attachment = Assert.Single(report.DocumentInfo.Attachments);
        Assert.Equal("data.xml", attachment.Name);
        Assert.Equal("data.xml", attachment.FileName);
        Assert.Equal("text/xml", attachment.MimeType);
        Assert.Equal(PdfAssociatedFileRelationship.Data, attachment.Relationship);
        Assert.Equal(4, attachment.SizeBytes);
        Assert.True(attachment.IsAssociatedFile);
        Assert.Same(attachment, Assert.Single(report.DocumentInfo.GetAttachmentsBySource("AF")));
        Assert.Same(attachment, Assert.Single(report.DocumentInfo.GetAttachmentsByRelationship(PdfAssociatedFileRelationship.Data)));
        Assert.Empty(report.DocumentInfo.GetAttachmentsBySource("Names/EmbeddedFiles"));
        Assert.DoesNotContain(report.RewriteBlockers, blocker => blocker.Kind == PdfRewriteBlockerKind.EmbeddedFiles);
    }

    [Fact]
    public void Preflight_AllowsSimpleOptionalContentPdfReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildOptionalContentPdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasOptionalContent);
        Assert.False(report.Probe.HasEmbeddedFiles);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasOptionalContent);
        Assert.True(report.DocumentInfo.HasReadableOptionalContent);
        Assert.True(report.DocumentInfo.HasOptionalContentGroups);
        Assert.Equal(1, report.DocumentInfo.OptionalContentGroupCount);
        Assert.Equal(new[] { "Layer 1" }, report.DocumentInfo.OptionalContentGroupNames);
        PdfOptionalContentGroup group = Assert.Single(report.DocumentInfo.OptionalContentGroups);
        Assert.Equal(5, group.ObjectNumber);
        Assert.Equal("Layer 1", group.Name);
        Assert.True(group.IsInitiallyVisible);
        Assert.False(group.IsLocked);
        Assert.True(group.IsInDefaultOrder);
        Assert.Same(group, Assert.Single(report.DocumentInfo.GetOptionalContentGroupsByName("Layer 1")));
        Assert.False(report.DocumentInfo.HasEmbeddedFiles);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.RewriteBlockers);
        Assert.DoesNotContain("PDF optional content layers are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
    }

    [Fact]
    public void Inspect_ReadsOptionalContentLayerMetadata() {
        PdfDocumentInfo info = PdfInspector.Inspect(PdfOptionalContentSupport.BuildOptionalContentMetadataPdf());

        Assert.True(info.HasOptionalContent);
        Assert.True(info.HasReadableOptionalContent);
        Assert.True(info.HasOptionalContentGroups);
        Assert.Equal(2, info.OptionalContentGroupCount);
        Assert.Equal(new[] { "Print layer", "Hidden layer" }, info.OptionalContentGroupNames);
        Assert.NotNull(info.OptionalContent);
        Assert.Equal("Default layers", info.OptionalContent!.DefaultConfigurationName);
        Assert.Equal("OfficeIMO fixture", info.OptionalContent.DefaultConfigurationCreator);
        Assert.Equal("ON", info.OptionalContent.BaseState);
        Assert.Equal(new[] { 5 }, info.OptionalContent.OnGroupObjectNumbers);
        Assert.Equal(new[] { 6 }, info.OptionalContent.OffGroupObjectNumbers);
        Assert.Equal(new[] { 6 }, info.OptionalContent.LockedGroupObjectNumbers);
        Assert.Equal(new[] { 5, 6 }, info.OptionalContent.OrderGroupObjectNumbers);

        PdfOptionalContentGroup printLayer = Assert.Single(info.GetOptionalContentGroupsByName("Print layer"));
        Assert.Equal(5, printLayer.ObjectNumber);
        Assert.Equal(new[] { "View", "Design" }, printLayer.Intents);
        Assert.True(printLayer.IsInitiallyVisible);
        Assert.False(printLayer.IsLocked);
        Assert.True(printLayer.IsInDefaultOrder);
        Assert.Equal("ON", printLayer.ViewState);
        Assert.Equal("ON", printLayer.PrintState);
        Assert.Equal("OFF", printLayer.ExportState);
        Assert.Equal("OfficeIMO", printLayer.UsageCreator);
        Assert.Equal("Artwork", printLayer.UsageSubtype);

        PdfOptionalContentGroup hiddenLayer = Assert.Single(info.GetOptionalContentGroupsByName("Hidden layer"));
        Assert.Equal(6, hiddenLayer.ObjectNumber);
        Assert.Equal(new[] { "View" }, hiddenLayer.Intents);
        Assert.False(hiddenLayer.IsInitiallyVisible);
        Assert.True(hiddenLayer.IsLocked);
        Assert.True(hiddenLayer.IsInDefaultOrder);
        Assert.Equal("OFF", hiddenLayer.ViewState);
        Assert.Equal("OFF", hiddenLayer.PrintState);
        Assert.Equal("ON", hiddenLayer.ExportState);
        Assert.Empty(info.GetOptionalContentGroupsByName("Missing"));
    }

    [Fact]
    public void Preflight_AllowsComplexOptionalContentPdfReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildComplexOptionalContentPdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasOptionalContent);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasOptionalContent);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("PDF optional content layers are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.OptionalContent, "PDF optional content layers are not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Preflight_AllowsActiveContentPdfReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildActiveContentPdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.CanExtractText);
        Assert.True(report.CanExtractImages);
        Assert.True(report.CanReadLogicalObjects);
        Assert.False(report.CanManipulatePages);
        Assert.False(report.CanFillSimpleFormFields);
        Assert.False(report.CanFlattenSimpleFormFields);
        Assert.False(report.CanFillAndFlattenSimpleFormFields);
        Assert.False(report.Can(PdfPreflightCapability.FillSimpleFormFields));
        Assert.False(report.Can(PdfPreflightCapability.FlattenSimpleFormFields));
        Assert.Contains(
            "PDF active content is not supported for form filling or flattening by OfficeIMO.Pdf yet.",
            report.GetCapabilityDiagnostics(PdfPreflightCapability.FillSimpleFormFields));
        Assert.True(report.Probe.HasActiveContent);
        Assert.True(report.Probe.HasCatalogNameTrees);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasActiveContent);
        Assert.True(report.DocumentInfo.HasCatalogNameTrees);
        Assert.True(report.DocumentInfo.HasCatalogActions);
        Assert.Equal(1, report.DocumentInfo.CatalogActionCount);
        PdfCatalogAction action = Assert.Single(report.DocumentInfo.CatalogActions);
        Assert.Equal("Open", action.Name);
        Assert.Equal("JavaScript", action.ActionType);
        Assert.Equal("Names/JavaScript", action.Source);
        Assert.Null(action.TriggerName);
        Assert.Equal(new[] { "Open" }, report.DocumentInfo.CatalogActionNames);
        Assert.Equal(new[] { "JavaScript" }, report.DocumentInfo.CatalogActionTypes);
        Assert.Equal(new[] { "Names/JavaScript" }, report.DocumentInfo.CatalogActionSources);
        Assert.Same(action, Assert.Single(report.DocumentInfo.GetCatalogActionsByActionType("JavaScript")));
        Assert.Same(action, Assert.Single(report.DocumentInfo.GetCatalogActionsBySource("Names/JavaScript")));
        Assert.Empty(report.DocumentInfo.GetCatalogActionsByActionType("Launch"));
        Assert.Empty(report.DocumentInfo.GetCatalogActionsBySource("OpenAction"));
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("PDF active content is not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.ActiveContent, "PDF active content is not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Inspect_ReadsCatalogOpenActionAndAdditionalActionMetadata() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildCatalogActiveActionSlotsPdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasActiveContent);
        Assert.True(report.Probe.HasOpenActions);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasActiveContent);
        Assert.True(report.DocumentInfo.HasOpenActions);
        Assert.True(report.DocumentInfo.HasCatalogActions);
        Assert.Equal(4, report.DocumentInfo.CatalogActionCount);
        Assert.Equal(new[] { "JavaScript", "Launch", "RichMedia", "SubmitForm" }, report.DocumentInfo.CatalogActionTypes.OrderBy(type => type).ToArray());
        Assert.Equal(new[] { "OpenAction", "AA" }, report.DocumentInfo.CatalogActionSources);
        Assert.Equal(new[] { "AA.DS", "AA.WC", "OpenAction", "OpenAction.Next" }, report.DocumentInfo.CatalogActionNames.OrderBy(name => name).ToArray());

        PdfCatalogAction openAction = Assert.Single(report.DocumentInfo.GetCatalogActionsByActionType("JavaScript"));
        Assert.Equal("OpenAction", openAction.Name);
        Assert.Equal("OpenAction", openAction.Source);
        Assert.Null(openAction.TriggerName);

        PdfCatalogAction nextAction = Assert.Single(report.DocumentInfo.GetCatalogActionsByActionType("RichMedia"));
        Assert.Equal("OpenAction.Next", nextAction.Name);
        Assert.Equal("OpenAction", nextAction.Source);

        PdfCatalogAction launchAction = Assert.Single(report.DocumentInfo.GetCatalogActionsByActionType("Launch"));
        Assert.Equal("AA.WC", launchAction.Name);
        Assert.Equal("AA", launchAction.Source);
        Assert.Equal("WC", launchAction.TriggerName);

        PdfCatalogAction submitAction = Assert.Single(report.DocumentInfo.GetCatalogActionsByActionType("SubmitForm"));
        Assert.Equal("AA.DS", submitAction.Name);
        Assert.Equal("AA", submitAction.Source);
        Assert.Equal("DS", submitAction.TriggerName);

        Assert.Equal(2, report.DocumentInfo.GetCatalogActionsBySource("OpenAction").Count);
        Assert.Equal(2, report.DocumentInfo.GetCatalogActionsBySource("AA").Count);
        Assert.Empty(report.DocumentInfo.GetCatalogActionsBySource("Names/JavaScript"));
        Assert.Empty(report.DocumentInfo.GetCatalogActionsByActionType("GoTo"));
        Assert.Empty(report.ReadBlockers);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.ActiveContent, "PDF active content is not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Reader_ExposesCatalogActionDiagnostics() {
        PdfDocument document = PdfDocument.Open(BuildCatalogActiveActionSlotsPdf());

        IReadOnlyList<PdfCatalogAction> actions = document.Read.CatalogActions();
        Assert.Equal(4, actions.Count);
        Assert.True(document.Read.TryCatalogActions().Succeeded);

        PdfCatalogAction openAction = Assert.Single(document.Read.CatalogActionsByActionType("JavaScript"));
        Assert.Equal("OpenAction", openAction.Name);
        Assert.Equal("OpenAction", openAction.Source);
        Assert.Null(openAction.TriggerName);

        PdfCatalogAction submitAction = Assert.Single(document.Read.CatalogActionsByActionType("SubmitForm"));
        Assert.Equal("AA.DS", submitAction.Name);
        Assert.Equal("AA", submitAction.Source);
        Assert.Equal("DS", submitAction.TriggerName);

        Assert.Equal(2, document.Read.CatalogActionsBySource("OpenAction").Count);
        Assert.Equal(2, document.Read.CatalogActionsBySource("AA").Count);
        Assert.Empty(document.Read.CatalogActionsBySource("Names/JavaScript"));
        Assert.True(document.Read.TryCatalogActionsByActionType("Launch").Succeeded);
        Assert.True(document.Read.TryCatalogActionsBySource("AA").Succeeded);
    }

    [Fact]
    public void Inspect_ReadsAnnotationActionMetadataWithoutTreatingItAsNavigationLinks() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildActiveAnnotationActionsPdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasActiveContent);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasActiveContent);
        Assert.Equal(2, report.DocumentInfo.AnnotationCount);
        Assert.Equal(3, report.DocumentInfo.AnnotationActionTypeCount);
        Assert.Equal(new[] { "JavaScript", "Launch", "SubmitForm" }, report.DocumentInfo.AnnotationActionTypes);
        Assert.Equal(0, report.DocumentInfo.LinkAnnotationCount);
        Assert.Empty(report.DocumentInfo.LinkAnnotations);

        PdfAnnotation scriptLink = Assert.Single(report.DocumentInfo.GetAnnotationsByActionType("JavaScript"));
        Assert.Equal("Link", scriptLink.Subtype);
        Assert.Equal("Script link", scriptLink.Contents);
        Assert.Equal("JavaScript", scriptLink.ActionType);
        Assert.True(scriptLink.HasAction);
        Assert.False(scriptLink.HasAdditionalActions);

        PdfAnnotation textAnnotation = Assert.Single(report.DocumentInfo.GetAnnotationsBySubtype("Text"));
        Assert.False(textAnnotation.HasAction);
        Assert.True(textAnnotation.HasAdditionalActions);
        Assert.Equal(2, textAnnotation.AdditionalActions.Count);
        Assert.Equal("E", textAnnotation.AdditionalActions[0].TriggerName);
        Assert.Equal("Launch", textAnnotation.AdditionalActions[0].ActionType);
        Assert.Equal("X", textAnnotation.AdditionalActions[1].TriggerName);
        Assert.Equal("SubmitForm", textAnnotation.AdditionalActions[1].ActionType);
        Assert.Same(textAnnotation, Assert.Single(report.DocumentInfo.GetAnnotationsByActionType("Launch")));
        Assert.Same(textAnnotation, Assert.Single(report.DocumentInfo.GetAnnotationsByActionType("SubmitForm")));
        Assert.Empty(report.DocumentInfo.GetAnnotationsByActionType("GoTo"));
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.ActiveContent, "PDF active content is not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Inspect_ReadsPageAdditionalActionMetadata() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildPageAdditionalActionsPdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasActiveContent);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasActiveContent);
        Assert.Equal(2, report.DocumentInfo.PageActionCount);
        Assert.True(report.DocumentInfo.HasPageActions);
        Assert.Equal(new[] { "JavaScript", "Launch" }, report.DocumentInfo.PageActionTypes);
        Assert.Equal(new[] { "O", "C" }, report.DocumentInfo.PageActionTriggerNames);
        Assert.Equal(0, report.DocumentInfo.LinkAnnotationCount);
        Assert.Empty(report.DocumentInfo.LinkAnnotations);
        Assert.Equal(0, report.DocumentInfo.AnnotationCount);
        Assert.Empty(report.DocumentInfo.Annotations);

        PdfPageInfo page = Assert.Single(report.DocumentInfo.Pages);
        Assert.True(page.HasPageActions);
        Assert.Equal(2, page.PageActionCount);
        Assert.Equal(1, page.PageActions[0].PageNumber);
        Assert.Equal("O", page.PageActions[0].TriggerName);
        Assert.Equal("O", page.PageActions[0].ActionPath);
        Assert.Equal("JavaScript", page.PageActions[0].ActionType);
        Assert.False(page.PageActions[0].IsChainedAction);
        Assert.Equal("C", page.PageActions[1].TriggerName);
        Assert.Equal("C", page.PageActions[1].ActionPath);
        Assert.Equal("Launch", page.PageActions[1].ActionType);

        PdfPageAction openAction = Assert.Single(report.DocumentInfo.GetPageActionsByTriggerName("O"));
        Assert.Same(openAction, Assert.Single(report.DocumentInfo.GetPageActionsByActionType("JavaScript")));
        Assert.Equal(1, openAction.PageNumber);
        Assert.Equal(2, report.DocumentInfo.GetPageActions(1).Count);
        Assert.Same(openAction, Assert.Single(report.DocumentInfo.GetPageActionsByActionPath("O")));
        Assert.Empty(report.DocumentInfo.GetPageActions(2));
        Assert.Empty(report.DocumentInfo.GetPageActionsByActionType("GoTo"));
        Assert.Empty(report.DocumentInfo.GetPageActionsByTriggerName("D"));
        Assert.Empty(report.DocumentInfo.GetPageActionsByActionPath("O.Next"));
        Assert.Equal(new[] { "O", "C" }, report.DocumentInfo.PageActionPaths);
        Assert.Equal(2, report.DocumentInfo.PageActionsByActionPath.Count);
        Assert.Empty(report.DocumentInfo.Pages[0].Annotations);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.ActiveContent, "PDF active content is not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Reader_ExposesPageActionDiagnostics() {
        PdfDocument document = PdfDocument.Open(BuildPageAdditionalActionsPdf());

        IReadOnlyList<PdfPageAction> actions = document.Read.PageActions();
        Assert.Equal(2, actions.Count);
        Assert.True(document.Read.TryPageActions().Succeeded);

        IReadOnlyList<PdfPageAction> pageActions = document.Read.PageActions(1);
        Assert.Equal(2, pageActions.Count);

        PdfPageAction openAction = Assert.Single(document.Read.PageActionsByTriggerName("O"));
        Assert.Equal(1, openAction.PageNumber);
        Assert.Equal("O", openAction.ActionPath);
        Assert.Equal("JavaScript", openAction.ActionType);
        Assert.False(openAction.IsChainedAction);

        PdfPageAction actionTypeMatch = Assert.Single(document.Read.PageActionsByActionType("JavaScript"));
        Assert.Equal(openAction.PageNumber, actionTypeMatch.PageNumber);
        Assert.Equal(openAction.TriggerName, actionTypeMatch.TriggerName);
        Assert.Equal(openAction.ActionPath, actionTypeMatch.ActionPath);

        PdfPageAction actionPathMatch = Assert.Single(document.Read.PageActionsByActionPath("O"));
        Assert.Equal(openAction.PageNumber, actionPathMatch.PageNumber);
        Assert.Equal(openAction.TriggerName, actionPathMatch.TriggerName);
        Assert.Equal(openAction.ActionType, actionPathMatch.ActionType);
        Assert.Empty(document.Read.PageActions(2));
        Assert.Empty(document.Read.PageActionsByActionType("GoTo"));
        Assert.Empty(document.Read.PageActionsByTriggerName("D"));
        Assert.Empty(document.Read.PageActionsByActionPath("O.Next"));
        Assert.True(document.Read.TryPageActions(1).Succeeded);
        Assert.True(document.Read.TryPageActionsByActionType("Launch").Succeeded);
        Assert.True(document.Read.TryPageActionsByTriggerName("C").Succeeded);
        Assert.True(document.Read.TryPageActionsByActionPath("C").Succeeded);
    }

    [Fact]
    public void Inspect_ReadsAnnotationNextActionMetadataWithoutPayloads() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildChainedAnnotationActionsPdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasActiveContent);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasActiveContent);
        Assert.Equal(2, report.DocumentInfo.AnnotationCount);
        Assert.Equal(5, report.DocumentInfo.AnnotationActionTypeCount);
        Assert.Equal(new[] { "JavaScript", "Launch", "RichMedia", "SubmitForm", "ImportData" }, report.DocumentInfo.AnnotationActionTypes);

        PdfAnnotation scriptLink = Assert.Single(report.DocumentInfo.GetAnnotationsByActionType("JavaScript"));
        Assert.True(scriptLink.HasAction);
        Assert.True(scriptLink.HasChainedActions);
        Assert.Equal(2, scriptLink.ChainedActionCount);
        Assert.Equal("A", scriptLink.ChainedActions[0].SourceName);
        Assert.Equal("A.Next.0", scriptLink.ChainedActions[0].ActionPath);
        Assert.Equal("Launch", scriptLink.ChainedActions[0].ActionType);
        Assert.Equal("A.Next.1", scriptLink.ChainedActions[1].ActionPath);
        Assert.Equal("RichMedia", scriptLink.ChainedActions[1].ActionType);
        Assert.Same(scriptLink, Assert.Single(report.DocumentInfo.GetAnnotationsByActionType("Launch")));
        Assert.Same(scriptLink, Assert.Single(report.DocumentInfo.GetAnnotationsByActionType("RichMedia")));

        PdfAnnotation note = Assert.Single(report.DocumentInfo.GetAnnotationsBySubtype("Text"));
        Assert.False(note.HasAction);
        Assert.True(note.HasAdditionalActions);
        Assert.True(note.HasChainedActions);
        Assert.Equal("E", Assert.Single(note.AdditionalActions).TriggerName);
        Assert.Equal("SubmitForm", Assert.Single(note.AdditionalActions).ActionType);
        PdfAnnotationChainedAction chainedAction = Assert.Single(note.ChainedActions);
        Assert.Equal("E", chainedAction.SourceName);
        Assert.Equal("E.Next", chainedAction.ActionPath);
        Assert.Equal("ImportData", chainedAction.ActionType);
        Assert.Same(note, Assert.Single(report.DocumentInfo.GetAnnotationsByActionType("ImportData")));
        Assert.Empty(report.DocumentInfo.GetAnnotationsByActionType("GoTo"));
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.ActiveContent, "PDF active content is not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Inspect_ReadsReusedIndirectAnnotationNextActionsForEachSource() {
        PdfDocumentInfo info = PdfInspector.Inspect(BuildSharedChainedAnnotationActionsPdf());

        PdfAnnotation annotation = Assert.Single(info.Annotations);
        Assert.Equal("JavaScript", annotation.ActionType);
        Assert.Equal(2, annotation.ChainedActionCount);
        Assert.Equal(new[] { "A", "E" }, annotation.ChainedActions.Select(action => action.SourceName).ToArray());
        Assert.Equal(new[] { "A.Next", "E.Next" }, annotation.ChainedActions.Select(action => action.ActionPath).ToArray());
        Assert.All(annotation.ChainedActions, action => Assert.Equal("Launch", action.ActionType));
        Assert.Same(annotation, Assert.Single(info.GetAnnotationsByActionType("Launch")));
    }

    [Fact]
    public void Inspect_ReadsPageNextActionMetadataWithoutPayloads() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildPageChainedActionsPdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasActiveContent);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasPageActions);
        Assert.Equal(3, report.DocumentInfo.PageActionCount);
        Assert.Equal(new[] { "JavaScript", "Launch", "RichMedia" }, report.DocumentInfo.PageActionTypes);
        Assert.Equal(new[] { "O" }, report.DocumentInfo.PageActionTriggerNames);
        Assert.Equal(new[] { "O", "O.Next.0", "O.Next.1" }, report.DocumentInfo.PageActionPaths);

        PdfPageInfo page = Assert.Single(report.DocumentInfo.Pages);
        Assert.Equal(3, page.PageActionCount);
        Assert.False(page.PageActions[0].IsChainedAction);
        Assert.True(page.PageActions[1].IsChainedAction);
        Assert.True(page.PageActions[2].IsChainedAction);
        Assert.All(page.PageActions, action => Assert.Equal("O", action.TriggerName));

        PdfPageAction launchAction = Assert.Single(report.DocumentInfo.GetPageActionsByActionPath("O.Next.0"));
        Assert.Equal("Launch", launchAction.ActionType);
        Assert.Same(launchAction, Assert.Single(report.DocumentInfo.GetPageActionsByActionType("Launch")));
        Assert.Equal(3, report.DocumentInfo.GetPageActionsByTriggerName("O").Count);
        Assert.Empty(report.DocumentInfo.GetPageActionsByActionPath("O.Next.2"));
        Assert.Empty(report.DocumentInfo.LinkAnnotations);
        Assert.Empty(report.DocumentInfo.Annotations);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.ActiveContent, "PDF active content is not supported for rewriting by OfficeIMO.Pdf yet.");
    }


}
