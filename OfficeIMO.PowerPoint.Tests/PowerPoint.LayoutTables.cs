using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointLayoutTableTests {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ContinuationsPreserveEveryDataRowAndTheRequestedHeaderPolicy(bool includeHeaders) {
        using var document = PowerPointPresentation.Create(new PowerPointCreateOptions { Theme = OfficeTheme.Modern });
        var table = new PowerPointLayoutTable(new[] { "Service", "Count" },
            Enumerable.Range(1, 7).Select(index => (IReadOnlyList<string>)new[] { "Service " + index, index.ToString() }).ToArray()) {
            IncludeHeaders = includeHeaders, MaximumRows = 3
        };
        var options = PowerPointCompositionOptions.FromDesign(PowerPointDeckDesign.FromBrand("#2563EB", "table-content"));
        options.ApplyTheme = false;
        document.Compose(new PowerPointDeckPlan().AddContent("Services", new[] { table }), options);
        using var saved = PowerPointPresentation.Load(new MemoryStream(document.ToBytes()));
        Assert.Equal(3, saved.Slides.Count);
        var tables = saved.Slides.SelectMany(slide => slide.Tables).ToArray();
        Assert.Equal(3, tables.Length);
        var values = new List<string>();
        foreach (var native in tables) {
            if (includeHeaders) Assert.Equal("Service", native.GetCell(0, 0).Text);
            for (int row = includeHeaders ? 1 : 0; row < native.Rows; row++) values.Add(native.GetCell(row, 0).Text);
        }
        Assert.Equal(Enumerable.Range(1, 7).Select(index => "Service " + index), values);
        Assert.Empty(saved.ValidateDocument());
    }
}
