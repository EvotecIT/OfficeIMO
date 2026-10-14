using System;
using System.Collections.Generic;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public class VisioRoutingAtomicityTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FailedRouteEnumerationRetainsExistingNativeOutput(bool nullEntry) {
        var document = VisioDocument.Create(); var page = document.AddPage("Atomic routes");
        var connector = page.AddConnector("route", new OfficePoint(1, 1), new OfficePoint(5, 3));
        connector.RouteThrough(new VisioConnectorWaypoint(2, 1), new VisioConnectorWaypoint(2, 3));
        byte[] before = document.ToLegacyXmlResult().Value;
        if (nullEntry) Assert.Throws<ArgumentException>(() => connector.RouteThrough(new[] { new VisioConnectorWaypoint(3, 1), null! }));
        else Assert.Throws<InvalidOperationException>(() => connector.RouteThrough(FailingEnumeration()));
        Assert.Equal(before, document.ToLegacyXmlResult().Value);
    }

    [Theory]
    [InlineData(double.NaN, 1)]
    [InlineData(1, double.PositiveInfinity)]
    public void RejectsNonFiniteRouteWithoutChangingExistingNativeOutput(double x, double y) {
        var document = VisioDocument.Create(); var page = document.AddPage("Finite routes");
        var connector = page.AddConnector("route", new OfficePoint(1, 1), new OfficePoint(5, 3));
        connector.RouteThrough(new VisioConnectorWaypoint(2, 1), new VisioConnectorWaypoint(2, 3));
        byte[] before = document.ToLegacyXmlResult().Value;
        Assert.Throws<ArgumentException>(() => connector.RouteThrough(new VisioConnectorWaypoint(x, y)));
        Assert.Equal(before, document.ToLegacyXmlResult().Value);
    }

    private static IEnumerable<VisioConnectorWaypoint> FailingEnumeration() {
        yield return new VisioConnectorWaypoint(3, 1);
        throw new InvalidOperationException("Replacement source failed.");
    }
}
