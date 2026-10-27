using System.Globalization;
using DocumentFormat.OpenXml.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;

namespace OfficeIMO.Word;

public partial class WordParagraph {
    /// <summary>Adds a legacy VML image to the paragraph.</summary>
    /// <param name="filePathImage">Image file to embed in the paragraph's document story.</param>
    /// <param name="width">Optional width in points. An omitted value leaves the VML width declaration absent.</param>
    /// <param name="height">Optional height in points. An omitted value leaves the VML height declaration absent.</param>
    /// <returns>The created image occurrence.</returns>
    public WordImage AddImageVml(string filePathImage, double? width = null, double? height = null) {
        using var stream = File.OpenRead(filePathImage);
        return InsertImageVml(stream, Path.GetFileName(filePathImage), width, height);
    }

    /// <summary>Embeds a VML image in the current run, returning its attached occurrence without disposing the caller's stream.</summary>
    /// <remarks>Dimensions are points. Omitted dimensions remain absent from the VML style; shared image embedding owns payload and story relationships.</remarks>
    internal WordImage InsertImageVml(Stream imageStream, string fileName, double? widthPoints = null,
        double? heightPoints = null, string description = "") {
        var run = VerifyRun();
        WordImageLocation location = WordImage.AddImageToLocation(_document, this, imageStream, fileName,
            widthPoints * 96D / 72D, heightPoints * 96D / 72D);
        string style = "mso-wrap-style:square";
        if (widthPoints.HasValue) style = "width:" + widthPoints.Value.ToString("R", CultureInfo.InvariantCulture) + "pt;" + style;
        if (heightPoints.HasValue) style = "height:" + heightPoints.Value.ToString("R", CultureInfo.InvariantCulture) + "pt;" + style;
        var shape = new V.Shape(new V.ImageData {
            RelationshipId = location.RelationshipId,
            Title = Path.GetFileName(fileName)
        }) {
            Id = "Image" + Guid.NewGuid().ToString("N"),
            Style = style,
            Type = "#_x0000_t75"
        };
        run.Append(new Picture(shape));
        var image = new WordImage(_document, _paragraph, run, shape) { Description = description };
        return image;
    }
}
