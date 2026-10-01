namespace OfficeIMO.Invoicing.Tests;

public sealed class InvoiceTargetCatalogTests {
    [Fact]
    public void CatalogTargetsAuthorReadableInvoicesWithTheRequestedProjectionPolicy() {
        var targets = InvoiceXmlOptions.GetSupportedTargets(InvoiceProjectionPolicy.AllowProfileDefinedDataLoss);
        Assert.Equal(10, targets.Count);
        Assert.Equal(targets.Count, targets.Select(target => (target.Release, target.Syntax, target.Profile)).Distinct().Count());
        foreach (var target in targets) {
            Assert.Equal(InvoiceProjectionPolicy.AllowProfileDefinedDataLoss, target.ProjectionPolicy);
            var read = InvoiceParser.Read(InvoiceSerializer.Write(InvoiceFixture.Create(), target));
            Assert.Equal(target.Syntax, read.Declaration.Syntax); Assert.Equal(target.Profile, read.Declaration.Profile);
            Assert.True(read.HasCompleteMapping);
        }
        Assert.All(InvoiceXmlOptions.GetSupportedTargets(), target => Assert.Equal(InvoiceProjectionPolicy.RejectDataLoss, target.ProjectionPolicy));
        Assert.Throws<ArgumentOutOfRangeException>(() => InvoiceXmlOptions.GetSupportedTargets((InvoiceProjectionPolicy)999));
    }
}
