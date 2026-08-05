using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointPowerShellValueTypeTests {
        [Fact]
        public void SlideLayoutTypeIsAClrEnumAndMapsToOpenXml() {
            Assert.True(typeof(PowerPointSlideLayoutType).IsEnum);
            Assert.Equal(SlideLayoutValues.Text, PowerPointSlideLayoutType.Text.ToOpenXml());
            Assert.Equal(SlideLayoutValues.MidiaAndText, PowerPointSlideLayoutType.MediaAndText.ToOpenXml());
            Assert.Equal(PowerPointSlideLayoutType.MediaAndText, SlideLayoutValues.MidiaAndText.ToOfficeIMO());

            PowerPointSlideLayoutType[] values = Enum.GetValues(typeof(PowerPointSlideLayoutType))
                .Cast<PowerPointSlideLayoutType>()
                .ToArray();
            Assert.Equal(values.Length, values.Select(value => value.ToOpenXml()).Distinct().Count());
            Assert.All(values, value => Assert.Equal(value, value.ToOpenXml().ToOfficeIMO()));
        }

        [Fact]
        public void SlideLayoutMethodsAcceptOfficeIMOType() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                int textLayout = presentation.GetLayoutIndexWithType(PowerPointSlideLayoutType.Text);
                PowerPointSlide slide = presentation.AddSlideWithLayoutType(PowerPointSlideLayoutType.Text);

                Assert.Equal(textLayout, slide.LayoutIndex);
                Assert.Contains(presentation.GetSlideLayouts(), layout => layout.LayoutType == PowerPointSlideLayoutType.Text);
                slide.SetLayoutWithType(PowerPointSlideLayoutType.Blank);
                Assert.Equal(presentation.GetLayoutIndexWithType(PowerPointSlideLayoutType.Blank), slide.LayoutIndex);
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }
    }
}
