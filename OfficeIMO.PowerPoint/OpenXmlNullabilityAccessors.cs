using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        private Presentation PresentationRoot {
            get => _presentationPart.Presentation ??= new Presentation();
            set => _presentationPart.Presentation = value;
        }
    }

    public partial class PowerPointSlide {
        private Slide SlideRoot {
            get => _slidePart.Slide ?? throw new InvalidOperationException("Slide is null.");
            set => _slidePart.Slide = value;
        }
    }

    public partial class PowerPointPicture {
        private Slide SlideRoot =>
            _slidePart.Slide ?? throw new InvalidOperationException("Slide is null.");
    }

    public partial class PowerPointSlide {
        private static ChartSpace GetChartSpaceRoot(ChartPart chartPart) =>
            chartPart.ChartSpace ?? throw new InvalidOperationException("ChartSpace is null.");
    }
}
