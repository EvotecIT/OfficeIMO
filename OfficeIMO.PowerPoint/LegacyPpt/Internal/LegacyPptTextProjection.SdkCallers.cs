using DocumentFormat.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>
    /// The SDK form of the text projection for the table projection (PowerPointDocument.LegacyPpt.Tables.cs), which
    /// still builds SDK elements. Remove it when that projection writes through <see cref="OfficeXmlNode"/>.
    /// </summary>
    internal static partial class LegacyPptTextProjection {
        // SDK caller in PowerPointDocument.LegacyPpt.Tables.cs
        internal static P.TextBody CreateTextBody(LegacyPptTextBody source,
            LegacyPptTextFrameProperties? frame,
            Func<LegacyPptInteraction, IReadOnlyList<OpenXmlElement>>? projectInteraction,
            Func<LegacyPptPictureBullet, string?>? projectPictureBullet = null) {
            ArgumentNullException.ThrowIfNull(source);
            var textBody = new P.TextBody();
            InteractionProjector? projector = projectInteraction == null
                ? null
                : (interaction, runProperties) => {
                    foreach (OpenXmlElement element in projectInteraction(interaction)) {
                        runProperties.Append(OfficeXmlNode.FromSdk(element));
                    }
                };
            WriteTextBody(OfficeXmlNode.FromSdk(textBody), source, frame, projector, projectPictureBullet);
            return textBody;
        }
    }
}
