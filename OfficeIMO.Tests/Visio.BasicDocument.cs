using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioBasicDocument {
        [Fact]
        public void CanCreateBasicVisioDocument() {
            VisioDocument document = new();
            VisioPage page = document.AddPage("Page1");
            VisioShape shape1 = new("1");
            VisioShape shape2 = new("2");
            page.Shapes.Add(shape1);
            page.Shapes.Add(shape2);
            page.Connectors.Add(new VisioConnector(shape1, shape2));

            Assert.Single(document.Pages);
            Assert.Equal(2, page.Shapes.Count);
            Assert.Single(page.Connectors);
        }
    }
}

