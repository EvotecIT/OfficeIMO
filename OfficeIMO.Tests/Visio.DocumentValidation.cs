using System;
using System.IO;
using System.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioDocumentValidation {
        [Fact]
        public void ValidateReturnsNoIssuesForSimpleDocument() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Page-1");
            page.Shapes.Add(new VisioShape("1", 1, 1, 2, 1, "Start"));

            Assert.Empty(document.Validate());
        }

        [Fact]
        public void ValidateReportsDuplicateIdsAndCrossPageConnectorIssues() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage first = document.AddPage("Page-1", id: 0);
            VisioPage second = document.AddPage("Page-2", id: 0);

            VisioShape shared = new("dup", 1, 1, 1, 1, "Shared");
            first.Shapes.Add(shared);
            first.Shapes.Add(new VisioShape("dup", 3, 1, 1, 1, "Duplicate"));

            VisioShape external = new("outside", 1, 1, 1, 1, "Outside");
            second.Shapes.Add(external);

            first.Connectors.Add(new VisioConnector("dup", shared, external));

            string[] issues = document.Validate().ToArray();

            Assert.Contains(issues, issue => issue.Contains("Duplicate page id '0'"));
            Assert.Contains(issues, issue => issue.Contains("Duplicate shape id 'dup'"));
            Assert.Contains(issues, issue => issue.Contains("Duplicate connector id 'dup'"));
            Assert.Contains(issues, issue => issue.Contains("references a target shape that is not part of the page"));
        }

        [Fact]
        public void ValidateReportsNegativeDimensionsAndDetachedConnectionPoints() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Page-1");
            page.Width = -1;
            page.Height = -2;

            VisioShape shape = new("shape", 1, 1, -1, -3, "Broken");
            page.Shapes.Add(shape);

            VisioShape target = new("target", 3, 1, 1, 1, "Target");
            page.Shapes.Add(target);

            VisioConnector connector = new("connector", shape, target) {
                FromConnectionPoint = new VisioConnectionPoint(0, 0, 0, 0)
            };
            page.Connectors.Add(connector);

            string[] issues = document.Validate().ToArray();

            Assert.Contains(issues, issue => issue.Contains("must have a positive width"));
            Assert.Contains(issues, issue => issue.Contains("must have a positive height"));
            Assert.Contains(issues, issue => issue.Contains("cannot have a negative width"));
            Assert.Contains(issues, issue => issue.Contains("cannot have a negative height"));
            Assert.Contains(issues, issue => issue.Contains("source connection point"));
        }
    }
}
