using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioQualityTests {
        [Fact]
        public void VisualQualityAnalyzerReportsObviousLayoutProblems() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Quality", 6, 4);
            VisioShape first = new VisioShape("first", 1.5, 1, 2, 1, "First");
            VisioShape second = new VisioShape("second", 2.1, 1, 2, 1, "Second");
            page.Shapes.Add(first);
            page.Shapes.Add(second);
            page.Shapes.Add(new VisioShape("outside", 6.4, 1, 1, 1, "Outside"));

            VisioShape source = new VisioShape("source", 0.7, 3, 0.8, 0.5, "Source");
            VisioShape obstacle = new VisioShape("obstacle", 3, 3, 0.9, 0.9, "Obstacle");
            VisioShape target = new VisioShape("target", 5.3, 3, 0.8, 0.5, "Target");
            page.Shapes.Add(source);
            page.Shapes.Add(obstacle);
            page.Shapes.Add(target);
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left)
                .PlaceLabelAt(6.4, 4.2, width: 1.2, height: 0.3);
            connector.Label = "outside label";

            var issues = page.AnalyzeVisualQuality(new VisioDiagramQualityOptions {
                RequireConnectorLabels = true
            });

            Assert.Contains(issues, issue => issue.Kind == "ShapeOverlap" && issue.ShapeId == first.Id && issue.OtherShapeId == second.Id);
            Assert.Contains(issues, issue => issue.Kind == "ShapeOutsidePage" && issue.ShapeId == "outside");
            Assert.Contains(issues, issue => issue.Kind == "ConnectorCrossesShape" && issue.ShapeId == obstacle.Id && issue.ConnectorId == connector.Id);
            Assert.Contains(issues, issue => issue.Kind == "ConnectorLabelOutsidePage" && issue.ConnectorId == connector.Id);
        }

        [Fact]
        public void VisualQualityAnalyzerReportsConnectorLabelShapeOverlaps() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("LabelShapes", 7, 5);
            VisioShape source = page.AddRectangle(1, 2, 0.8, 0.5, "Source");
            VisioShape obstacle = page.AddRectangle(3, 2, 1, 1, "Obstacle");
            VisioShape target = page.AddRectangle(5, 2, 0.8, 0.5, "Target");
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left)
                .PlaceLabelAt(3, 2, width: 1.2, height: 0.4);
            connector.Label = "covers node";

            IReadOnlyList<VisioDiagramQualityIssue> issues = page.AnalyzeVisualQuality(new VisioDiagramQualityOptions {
                CheckShapeOverlaps = false,
                CheckConnectorShapeIntersections = false
            });

            Assert.Contains(issues, issue =>
                issue.Kind == "ConnectorLabelOverlapsShape" &&
                issue.ShapeId == obstacle.Id &&
                issue.ConnectorId == connector.Id);
        }

        [Fact]
        public void VisualQualityAnalyzerIgnoresBackgroundSurfaceLabelOverlaps() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("BackgroundLabel", 7, 5);
            VisioShape background = page.AddRectangle(3, 2, 4, 2, "Region");
            background.SetUserCell("OfficeIMO.Kind", "BackgroundSurface", "STR");
            VisioShape source = page.AddRectangle(1, 2, 0.8, 0.5, "Source");
            VisioShape target = page.AddRectangle(5, 2, 0.8, 0.5, "Target");
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left)
                .PlaceLabelAt(3, 2, width: 1.2, height: 0.4);
            connector.Label = "inside region";

            IReadOnlyList<VisioDiagramQualityIssue> issues = page.AnalyzeVisualQuality(new VisioDiagramQualityOptions {
                CheckShapeOverlaps = false,
                CheckConnectorShapeIntersections = false
            });

            Assert.DoesNotContain(issues, issue =>
                issue.Kind == "ConnectorLabelOverlapsShape" &&
                issue.ShapeId == background.Id &&
                issue.ConnectorId == connector.Id);
        }

        [Fact]
        public void VisualQualityAnalyzerIgnoresGeneratedBackgroundCaptionAdornments() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("BackgroundCaption", 7, 5);
            VisioShape background = new("zone", 3, 2, 4, 2, string.Empty) { NameU = "Rectangle" };
            page.Shapes.Add(background);
            background.SetUserCell("OfficeIMO.Kind", "BackgroundSurface", "STR");
            VisioShape caption = page.AddTextBox("zone-label", 3, 2.85, 3.6, 0.3, "Zone");
            VisioShape source = page.AddRectangle(0.45, 2.85, 0.6, 0.4, "Source");
            VisioShape target = page.AddRectangle(5.55, 2.85, 0.6, 0.4, "Target");
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left);

            IReadOnlyList<VisioDiagramQualityIssue> issues = page.AnalyzeVisualQuality();

            Assert.True(caption.IsDiagramAdornment);
            Assert.DoesNotContain(issues, issue => issue.Kind == "ShapeOverlap" && (issue.ShapeId == background.Id || issue.OtherShapeId == background.Id));
            Assert.DoesNotContain(issues, issue => issue.Kind == "ConnectorCrossesShape" && issue.ShapeId == caption.Id && issue.ConnectorId == connector.Id);
        }

        [Fact]
        public void VisualQualityAnalyzerIgnoresCalloutsPlacedOnBackgroundSurfaces() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("BackgroundCallout", 7, 5);
            VisioShape background = new("lane", 3, 2.5, 5, 2.5, string.Empty) { NameU = "Rectangle" };
            page.Shapes.Add(background);
            background.SetUserCell("OfficeIMO.Kind", "BackgroundSurface", "STR");
            VisioShape target = page.AddRectangle(3, 2.4, 0.8, 0.5, "Target");
            VisioShape callout = page.AddCallout(target, "target-note", "Evidence note", VisioSide.Bottom);

            IReadOnlyList<VisioDiagramQualityIssue> issues = page.AnalyzeVisualQuality();

            Assert.True(callout.IsCallout);
            Assert.DoesNotContain(issues, issue => issue.Kind == "ShapeOverlap" && (issue.ShapeId == background.Id || issue.OtherShapeId == background.Id));
        }

        [Fact]
        public void VisualQualityAnalyzerReportsConnectorLabelOverlaps() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("LabelLabels", 7, 5);
            VisioShape left = page.AddRectangle(1, 2, 0.8, 0.5, "Left");
            VisioShape middle = page.AddRectangle(3, 2, 0.8, 0.5, "Middle");
            VisioShape right = page.AddRectangle(5, 2, 0.8, 0.5, "Right");
            VisioConnector first = page.AddConnector(left, middle, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left)
                .PlaceLabelAt(3, 3.4, width: 1.2, height: 0.4);
            first.Label = "first";
            VisioConnector second = page.AddConnector(middle, right, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left)
                .PlaceLabelAt(3.2, 3.4, width: 1.2, height: 0.4);
            second.Label = "second";

            IReadOnlyList<VisioDiagramQualityIssue> issues = page.AnalyzeVisualQuality(new VisioDiagramQualityOptions {
                CheckShapeOverlaps = false,
                CheckConnectorShapeIntersections = false,
                CheckConnectorLabelShapeOverlaps = false
            });

            Assert.Contains(issues, issue =>
                issue.Kind == "ConnectorLabelOverlap" &&
                issue.ConnectorId == first.Id &&
                issue.OtherConnectorId == second.Id);
        }

        [Fact]
        public void VisualQualityAnalyzerIgnoresContainerLikeOverlapsByDefault() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Containers", 7, 5);
            page.Shapes.Add(new VisioShape("container", 3.5, 2.5, 5, 3, "Container"));
            page.Shapes.Add(new VisioShape("child", 3.5, 2.5, 1, 0.7, "Child"));

            var issues = page.AnalyzeVisualQuality();

            Assert.DoesNotContain(issues, issue => issue.Kind == "ShapeOverlap");
        }

        [Fact]
        public void VisualQualityReportSummarizesIssuesAndQualityGateThrowsAtRequestedSeverity() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Gate", 6, 4);
            page.Shapes.Add(new VisioShape("first", 1.5, 1, 2, 1, "First"));
            page.Shapes.Add(new VisioShape("second", 2.1, 1, 2, 1, "Second"));
            page.Shapes.Add(new VisioShape("outside", 6.4, 1, 1, 1, "Outside"));

            VisioShape source = page.AddRectangle(1, 3, 0.8, 0.5, "Source");
            VisioShape target = page.AddRectangle(5, 3, 0.8, 0.5, "Target");
            page.AddConnector(source, target, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left);

            VisioDiagramQualityReport report = document.GetVisualQualityReport(new VisioDiagramQualityOptions {
                RequireConnectorLabels = true
            });

            Assert.Equal(1, report.ErrorCount);
            Assert.True(report.WarningCount >= 1);
            Assert.Equal(1, report.InformationCount);
            Assert.False(report.IsClean);
            Assert.Contains("ShapeOutsidePage", report.ToString());

            VisioDiagramQualityException exception = Assert.Throws<VisioDiagramQualityException>(() =>
                document.EnsureVisualQuality(new VisioDiagramQualityOptions {
                    RequireConnectorLabels = true
                }));

            Assert.Equal(VisioDiagramQualityIssueSeverity.Warning, exception.MinimumSeverity);
            Assert.DoesNotContain(exception.Issues, issue => issue.Severity == VisioDiagramQualityIssueSeverity.Information);
            Assert.Contains("quality gate failed", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void VisualQualityGateCanTreatInformationAsOptionalOrBlocking() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("InfoOnly", 6, 4);
            VisioShape source = page.AddRectangle(1, 2, 0.8, 0.5, "Source");
            VisioShape target = page.AddRectangle(5, 2, 0.8, 0.5, "Target");
            page.AddConnector(source, target, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left);

            VisioDiagramQualityOptions options = new VisioDiagramQualityOptions {
                RequireConnectorLabels = true,
                CheckConnectorShapeIntersections = false
            };

            VisioDiagramQualityReport report = page.GetVisualQualityReport(options);

            Assert.Equal(1, report.InformationCount);
            Assert.True(report.IsClean);
            Assert.Same(page, page.EnsureVisualQuality(options));

            VisioDiagramQualityException exception = Assert.Throws<VisioDiagramQualityException>(() =>
                page.EnsureVisualQuality(options, VisioDiagramQualityIssueSeverity.Information));

            Assert.Single(exception.Issues);
            Assert.Equal("ConnectorMissingLabel", exception.Issues[0].Kind);
        }

        [Fact]
        public void GallerySamplesGenerateValidAndVisuallyCleanDocuments() {
            string folderPath = Path.Combine(Path.GetTempPath(), "OfficeIMO-Visio-Gallery-" + Guid.NewGuid());

            IReadOnlyList<VisioGalleryResult> results = VisioGallery.Create(folderPath);

            Assert.Equal(13, results.Count);
            Assert.All(results, result => Assert.True(File.Exists(result.FilePath), result.FilePath));
            Assert.All(results, result => Assert.Null(result.DesktopValidation));
            Assert.All(results, result => Assert.Empty(result.PackageIssues));
            Assert.All(results, result => Assert.Empty(result.QualityIssues.Select(issue => issue.ToString())));

            VisioGalleryResult inventory = results.Single(result => result.Name == "CI/CD Inventory Graph");
            VisioDocument loaded = VisioDocument.Load(inventory.FilePath);
            VisioPage page = Assert.Single(loaded.Pages);
            Assert.Contains(page.Shapes, shape => shape.Id == "delivery-cluster" && shape.GetShapeDataValue("Owner") == "DevEx");
            Assert.Contains(page.Shapes, shape => shape.Id == "runtime-cluster" && shape.GetShapeDataValue("Owner") == "SRE");
            Assert.Contains(page.Shapes, shape => shape.Id == "legend-title" && shape.Text == "Legend");
            Assert.Contains(page.Connectors, connector => connector.Id == "agent-data-registry" && connector.GetShapeDataValue("Protocol") == "OCI");
            VisioStencilProfile profile = loaded.CreateStencilProfile();
            Assert.Contains("Containers and Kubernetes", profile.StencilCatalogs);
            Assert.Contains("Cloud", profile.StencilCatalogs);
            Assert.Contains("Infrastructure", profile.StencilCatalogs);

            VisioGalleryResult identity = results.Single(result => result.Name == "Identity Authentication Graph");
            VisioDocument loadedIdentity = VisioDocument.Load(identity.FilePath);
            VisioPage identityPage = Assert.Single(loadedIdentity.Pages);
            Assert.Contains(identityPage.Shapes, shape => shape.Id == "trust-boundary" && shape.GetShapeDataValue("Owner") == "Identity Security");
            Assert.Contains(identityPage.Shapes, shape => shape.Id == "legend-title" && shape.Text == "Legend");
            Assert.Contains(identityPage.Connectors, connector => connector.Id == "idp-data-app" && connector.GetShapeDataValue("Lifetime") == "60 minutes");
            VisioStencilProfile identityProfile = loadedIdentity.CreateStencilProfile();
            Assert.Contains("Security and Identity", identityProfile.StencilCatalogs);
            Assert.Contains("Collaboration and Business Process", identityProfile.StencilCatalogs);

            VisioGalleryResult kubernetes = results.Single(result => result.Name == "Kubernetes Service Mesh Graph");
            VisioDocument loadedKubernetes = VisioDocument.Load(kubernetes.FilePath);
            VisioPage kubernetesPage = Assert.Single(loadedKubernetes.Pages);
            Assert.Contains(kubernetesPage.Shapes, shape => shape.Id == "mesh-cluster" && shape.GetShapeDataValue("Owner") == "Platform Mesh");
            Assert.Contains(kubernetesPage.Shapes, shape => shape.Id == "legend-title" && shape.Text == "Legend");
            Assert.Contains(kubernetesPage.Connectors, connector => connector.Id == "api-data-stream" && connector.GetShapeDataValue("Format") == "CloudEvents");
            VisioStencilProfile kubernetesProfile = loadedKubernetes.CreateStencilProfile();
            Assert.Contains("Containers and Kubernetes", kubernetesProfile.StencilCatalogs);
            Assert.Contains("Data and Platform", kubernetesProfile.StencilCatalogs);

            VisioGalleryResult application = results.Single(result => result.Name == "Application Dependency Graph");
            VisioDocument loadedApplication = VisioDocument.Load(application.FilePath);
            VisioPage applicationPage = Assert.Single(loadedApplication.Pages);
            Assert.Contains(applicationPage.Shapes, shape => shape.Id == "runtime-cluster" && shape.GetShapeDataValue("Owner") == "Digital Platform");
            Assert.Contains(applicationPage.Shapes, shape => shape.Id == "legend-title" && shape.Text == "Legend");
            Assert.Contains(applicationPage.Connectors, connector => connector.Id == "api-data-sql" && connector.GetShapeDataValue("Protocol") == "SQL");
            VisioStencilProfile applicationProfile = loadedApplication.CreateStencilProfile();
            Assert.Contains("Cloud", applicationProfile.StencilCatalogs);
            Assert.Contains("Data and Platform", applicationProfile.StencilCatalogs);
            Assert.Contains("Security and Identity", applicationProfile.StencilCatalogs);

            VisioGalleryResult incident = results.Single(result => result.Name == "Incident Runbook Sequence");
            VisioDocument loadedIncident = VisioDocument.Load(incident.FilePath);
            VisioPage incidentPage = Assert.Single(loadedIncident.Pages);
            Assert.Contains(incidentPage.Shapes, shape => shape.Id == "recovery-fragment" && shape.GetUserCellValue("OfficeIMO.Kind") == "SequenceFragment");
            Assert.Contains(incidentPage.Shapes, shape => shape.Id == "runbook-note" && shape.GetUserCellValue("OfficeIMO.SequenceParticipantId") == "runbook");
            Assert.Contains(incidentPage.Shapes, shape => shape.Id == "api-active" && shape.GetUserCellValue("OfficeIMO.SequenceParticipantId") == "api");
            Assert.Contains(incidentPage.Connectors, connector => connector.Id == "record" && connector.Waypoints.Count == 2);
            VisioStencilProfile incidentProfile = loadedIncident.CreateStencilProfile();
            Assert.Contains("Sequence Diagram", incidentProfile.StencilCatalogs);
            Assert.Contains("SequenceActivation", incidentProfile.SemanticKinds);
            Assert.Contains("SequenceFragment", incidentProfile.SemanticKinds);
        }

        [Fact]
        public void GalleryDesktopValidationCanBeOptionalWhenVisioIsMissing() {
            if (VisioDesktopValidator.IsAvailable()) {
                return;
            }

            string optionalFolderPath = Path.Combine(Path.GetTempPath(), "OfficeIMO-Visio-Gallery-OptionalDesktop-" + Guid.NewGuid());
            IReadOnlyList<VisioGalleryResult> optionalResults = VisioGallery.Create(optionalFolderPath, new VisioGalleryOptions {
                ValidatePackage = false,
                AnalyzeVisualQuality = false,
                ValidateWithVisioDesktop = true
            });

            Assert.All(optionalResults, result => Assert.NotNull(result.DesktopValidation));
            Assert.All(optionalResults, result => Assert.False(result.DesktopValidation!.IsAvailable));
            Assert.All(optionalResults, result => Assert.True(result.IsClean));

            string strictFolderPath = Path.Combine(Path.GetTempPath(), "OfficeIMO-Visio-Gallery-StrictDesktop-" + Guid.NewGuid());
            IReadOnlyList<VisioGalleryResult> strictResults = VisioGallery.Create(strictFolderPath, new VisioGalleryOptions {
                ValidatePackage = false,
                AnalyzeVisualQuality = false,
                ValidateWithVisioDesktop = true,
                RequireVisioDesktop = true
            });

            Assert.All(strictResults, result => Assert.NotNull(result.DesktopValidation));
            Assert.All(strictResults, result => Assert.False(result.DesktopValidation!.IsAvailable));
            Assert.All(strictResults, result => Assert.False(result.IsClean));
        }
    }
}
