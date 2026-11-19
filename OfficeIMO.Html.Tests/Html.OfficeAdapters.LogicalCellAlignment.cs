using OfficeIMO.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class HtmlOfficeAdapters {
    [Theory]
    [InlineData("<td style='text-align:start'>Cell</td>", PowerPointTextAlignment.Left)]
    [InlineData("<td dir='rtl' style='text-align:start'>Cell</td>", PowerPointTextAlignment.Right)]
    [InlineData("<td dir='rtl' style='text-align:end'>Cell</td>", PowerPointTextAlignment.Left)]
    [InlineData("<td style='text-align:end'>Cell</td>", PowerPointTextAlignment.Right)]
    public void HtmlToPowerPoint_ResolvesLogicalCellAlignmentAgainstDirection(string cell, PowerPointTextAlignment expected) {
        using PowerPointDocument presentation = HtmlConversionDocument.Parse("<section class='officeimo-slide'><table><tr>" + cell + "</tr></table></section>").ToPowerPoint();
        PowerPointTableCell converted = presentation.Slides.SelectMany(slide => slide.Tables).Single().GetCell(0, 0);

        Assert.Equal(expected, converted.HorizontalAlignment);
    }
}
