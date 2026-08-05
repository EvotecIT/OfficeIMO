using DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint;

/// <summary>Specifies a built-in PowerPoint slide layout.</summary>
public enum PowerPointSlideLayoutType {
    /// <summary>Title slide.</summary>
    Title,
    /// <summary>Title and text.</summary>
    Text,
    /// <summary>Two-column text.</summary>
    TwoColumnText,
    /// <summary>Title and table.</summary>
    Table,
    /// <summary>Text and chart.</summary>
    TextAndChart,
    /// <summary>Chart and text.</summary>
    ChartAndText,
    /// <summary>Diagram.</summary>
    Diagram,
    /// <summary>Chart.</summary>
    Chart,
    /// <summary>Text and clip art.</summary>
    TextAndClipArt,
    /// <summary>Clip art and text.</summary>
    ClipArtAndText,
    /// <summary>Title only.</summary>
    TitleOnly,
    /// <summary>Blank slide.</summary>
    Blank,
    /// <summary>Text and object.</summary>
    TextAndObject,
    /// <summary>Object and text.</summary>
    ObjectAndText,
    /// <summary>Object only.</summary>
    ObjectOnly,
    /// <summary>Object layout.</summary>
    Object,
    /// <summary>Text and media.</summary>
    TextAndMedia,
    /// <summary>Media and text.</summary>
    MediaAndText,
    /// <summary>Object over text.</summary>
    ObjectOverText,
    /// <summary>Text over object.</summary>
    TextOverObject,
    /// <summary>Text and two objects.</summary>
    TextAndTwoObjects,
    /// <summary>Two objects and text.</summary>
    TwoObjectsAndText,
    /// <summary>Two objects over text.</summary>
    TwoObjectsOverText,
    /// <summary>Four objects.</summary>
    FourObjects,
    /// <summary>Vertical text.</summary>
    VerticalText,
    /// <summary>Clip art and vertical text.</summary>
    ClipArtAndVerticalText,
    /// <summary>Vertical title and text.</summary>
    VerticalTitleAndText,
    /// <summary>Vertical title and text over chart.</summary>
    VerticalTitleAndTextOverChart,
    /// <summary>Two objects.</summary>
    TwoObjects,
    /// <summary>Object and two objects.</summary>
    ObjectAndTwoObjects,
    /// <summary>Two objects and object.</summary>
    TwoObjectsAndObject,
    /// <summary>Custom layout.</summary>
    Custom,
    /// <summary>Section header.</summary>
    SectionHeader,
    /// <summary>Two text areas and two objects.</summary>
    TwoTextAndTwoObjects,
    /// <summary>Object with text.</summary>
    ObjectText,
    /// <summary>Picture with text.</summary>
    PictureText
}

internal static class PowerPointSlideLayoutTypeExtensions {
    internal static PowerPointSlideLayoutType ToOfficeIMO(this SlideLayoutValues value) => value switch {
        _ when value == SlideLayoutValues.Title => PowerPointSlideLayoutType.Title,
        _ when value == SlideLayoutValues.Text => PowerPointSlideLayoutType.Text,
        _ when value == SlideLayoutValues.TwoColumnText => PowerPointSlideLayoutType.TwoColumnText,
        _ when value == SlideLayoutValues.Table => PowerPointSlideLayoutType.Table,
        _ when value == SlideLayoutValues.TextAndChart => PowerPointSlideLayoutType.TextAndChart,
        _ when value == SlideLayoutValues.ChartAndText => PowerPointSlideLayoutType.ChartAndText,
        _ when value == SlideLayoutValues.Diagram => PowerPointSlideLayoutType.Diagram,
        _ when value == SlideLayoutValues.Chart => PowerPointSlideLayoutType.Chart,
        _ when value == SlideLayoutValues.TextAndClipArt => PowerPointSlideLayoutType.TextAndClipArt,
        _ when value == SlideLayoutValues.ClipArtAndText => PowerPointSlideLayoutType.ClipArtAndText,
        _ when value == SlideLayoutValues.TitleOnly => PowerPointSlideLayoutType.TitleOnly,
        _ when value == SlideLayoutValues.Blank => PowerPointSlideLayoutType.Blank,
        _ when value == SlideLayoutValues.TextAndObject => PowerPointSlideLayoutType.TextAndObject,
        _ when value == SlideLayoutValues.ObjectAndText => PowerPointSlideLayoutType.ObjectAndText,
        _ when value == SlideLayoutValues.ObjectOnly => PowerPointSlideLayoutType.ObjectOnly,
        _ when value == SlideLayoutValues.Object => PowerPointSlideLayoutType.Object,
        _ when value == SlideLayoutValues.TextAndMedia => PowerPointSlideLayoutType.TextAndMedia,
        _ when value == SlideLayoutValues.MidiaAndText => PowerPointSlideLayoutType.MediaAndText,
        _ when value == SlideLayoutValues.ObjectOverText => PowerPointSlideLayoutType.ObjectOverText,
        _ when value == SlideLayoutValues.TextOverObject => PowerPointSlideLayoutType.TextOverObject,
        _ when value == SlideLayoutValues.TextAndTwoObjects => PowerPointSlideLayoutType.TextAndTwoObjects,
        _ when value == SlideLayoutValues.TwoObjectsAndText => PowerPointSlideLayoutType.TwoObjectsAndText,
        _ when value == SlideLayoutValues.TwoObjectsOverText => PowerPointSlideLayoutType.TwoObjectsOverText,
        _ when value == SlideLayoutValues.FourObjects => PowerPointSlideLayoutType.FourObjects,
        _ when value == SlideLayoutValues.VerticalText => PowerPointSlideLayoutType.VerticalText,
        _ when value == SlideLayoutValues.ClipArtAndVerticalText => PowerPointSlideLayoutType.ClipArtAndVerticalText,
        _ when value == SlideLayoutValues.VerticalTitleAndText => PowerPointSlideLayoutType.VerticalTitleAndText,
        _ when value == SlideLayoutValues.VerticalTitleAndTextOverChart => PowerPointSlideLayoutType.VerticalTitleAndTextOverChart,
        _ when value == SlideLayoutValues.TwoObjects => PowerPointSlideLayoutType.TwoObjects,
        _ when value == SlideLayoutValues.ObjectAndTwoObjects => PowerPointSlideLayoutType.ObjectAndTwoObjects,
        _ when value == SlideLayoutValues.TwoObjectsAndObject => PowerPointSlideLayoutType.TwoObjectsAndObject,
        _ when value == SlideLayoutValues.Custom => PowerPointSlideLayoutType.Custom,
        _ when value == SlideLayoutValues.SectionHeader => PowerPointSlideLayoutType.SectionHeader,
        _ when value == SlideLayoutValues.TwoTextAndTwoObjects => PowerPointSlideLayoutType.TwoTextAndTwoObjects,
        _ when value == SlideLayoutValues.ObjectText => PowerPointSlideLayoutType.ObjectText,
        _ when value == SlideLayoutValues.PictureText => PowerPointSlideLayoutType.PictureText,
        _ => throw new ArgumentOutOfRangeException(nameof(value), value, "Unsupported Open XML slide layout type.")
    };

    internal static SlideLayoutValues ToOpenXml(this PowerPointSlideLayoutType value) => value switch {
        PowerPointSlideLayoutType.Title => SlideLayoutValues.Title,
        PowerPointSlideLayoutType.Text => SlideLayoutValues.Text,
        PowerPointSlideLayoutType.TwoColumnText => SlideLayoutValues.TwoColumnText,
        PowerPointSlideLayoutType.Table => SlideLayoutValues.Table,
        PowerPointSlideLayoutType.TextAndChart => SlideLayoutValues.TextAndChart,
        PowerPointSlideLayoutType.ChartAndText => SlideLayoutValues.ChartAndText,
        PowerPointSlideLayoutType.Diagram => SlideLayoutValues.Diagram,
        PowerPointSlideLayoutType.Chart => SlideLayoutValues.Chart,
        PowerPointSlideLayoutType.TextAndClipArt => SlideLayoutValues.TextAndClipArt,
        PowerPointSlideLayoutType.ClipArtAndText => SlideLayoutValues.ClipArtAndText,
        PowerPointSlideLayoutType.TitleOnly => SlideLayoutValues.TitleOnly,
        PowerPointSlideLayoutType.Blank => SlideLayoutValues.Blank,
        PowerPointSlideLayoutType.TextAndObject => SlideLayoutValues.TextAndObject,
        PowerPointSlideLayoutType.ObjectAndText => SlideLayoutValues.ObjectAndText,
        PowerPointSlideLayoutType.ObjectOnly => SlideLayoutValues.ObjectOnly,
        PowerPointSlideLayoutType.Object => SlideLayoutValues.Object,
        PowerPointSlideLayoutType.TextAndMedia => SlideLayoutValues.TextAndMedia,
        PowerPointSlideLayoutType.MediaAndText => SlideLayoutValues.MidiaAndText,
        PowerPointSlideLayoutType.ObjectOverText => SlideLayoutValues.ObjectOverText,
        PowerPointSlideLayoutType.TextOverObject => SlideLayoutValues.TextOverObject,
        PowerPointSlideLayoutType.TextAndTwoObjects => SlideLayoutValues.TextAndTwoObjects,
        PowerPointSlideLayoutType.TwoObjectsAndText => SlideLayoutValues.TwoObjectsAndText,
        PowerPointSlideLayoutType.TwoObjectsOverText => SlideLayoutValues.TwoObjectsOverText,
        PowerPointSlideLayoutType.FourObjects => SlideLayoutValues.FourObjects,
        PowerPointSlideLayoutType.VerticalText => SlideLayoutValues.VerticalText,
        PowerPointSlideLayoutType.ClipArtAndVerticalText => SlideLayoutValues.ClipArtAndVerticalText,
        PowerPointSlideLayoutType.VerticalTitleAndText => SlideLayoutValues.VerticalTitleAndText,
        PowerPointSlideLayoutType.VerticalTitleAndTextOverChart => SlideLayoutValues.VerticalTitleAndTextOverChart,
        PowerPointSlideLayoutType.TwoObjects => SlideLayoutValues.TwoObjects,
        PowerPointSlideLayoutType.ObjectAndTwoObjects => SlideLayoutValues.ObjectAndTwoObjects,
        PowerPointSlideLayoutType.TwoObjectsAndObject => SlideLayoutValues.TwoObjectsAndObject,
        PowerPointSlideLayoutType.Custom => SlideLayoutValues.Custom,
        PowerPointSlideLayoutType.SectionHeader => SlideLayoutValues.SectionHeader,
        PowerPointSlideLayoutType.TwoTextAndTwoObjects => SlideLayoutValues.TwoTextAndTwoObjects,
        PowerPointSlideLayoutType.ObjectText => SlideLayoutValues.ObjectText,
        PowerPointSlideLayoutType.PictureText => SlideLayoutValues.PictureText,
        _ => throw new ArgumentOutOfRangeException(nameof(value), value, "Unsupported PowerPoint slide layout type.")
    };
}
