using System.Security.Cryptography;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>
    /// SDK forms of the text projection for callers that still hold SDK elements: the table projection
    /// (PowerPointDocument.LegacyPpt.Tables.cs) and the readers that fingerprint projected text
    /// (LegacyPptProjectionMap.cs, LegacyPptShapeProjection.cs, LegacyPptPreservingWriter.cs). Remove each member when
    /// its callers read and write through <see cref="OfficeXmlNode"/>.
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

        // SDK caller in LegacyPptProjectionMap.cs and LegacyPptPreservingWriter.cs
        internal static string CreateFormattingFingerprint(
            P.TextBody? textBody, OpenXmlPart? ownerPart = null) {
            if (textBody == null) return string.Empty;
            var clone = (P.TextBody)textBody.CloneNode(true);
            if (clone.BodyProperties != null) {
                clone.ReplaceChild(new A.BodyProperties(),
                    clone.BodyProperties);
            }
            foreach (A.Text text in clone.Descendants<A.Text>()) text.Text = string.Empty;
            foreach (A.RunProperties properties in clone.Descendants<A.RunProperties>()
                         .ToArray()) {
                properties.RemoveAllChildren<A.HyperlinkOnClick>();
                properties.RemoveAllChildren<A.HyperlinkOnMouseOver>();
                if (!properties.HasAttributes && !properties.HasChildren) properties.Remove();
            }
            return clone.OuterXml + CreatePictureBulletImageFingerprint(
                textBody, ownerPart);
        }

        // SDK caller in LegacyPptProjectionMap.cs and LegacyPptPreservingWriter.cs (through CreateFormattingFingerprint)
        internal static string CreatePictureBulletImageFingerprint(
            OpenXmlElement? root, OpenXmlPart? ownerPart) {
            if (root == null || ownerPart == null) return string.Empty;
            var result = new System.Text.StringBuilder();
            foreach (A.PictureBullet bullet in root
                         .Descendants<A.PictureBullet>()) {
                string? relationshipId = bullet.GetFirstChild<A.Blip>()?
                    .Embed?.Value;
                result.Append("|buBlip:").Append(relationshipId);
                IdPartPair? pair = ownerPart.Parts.FirstOrDefault(item =>
                    string.Equals(item.RelationshipId, relationshipId,
                        StringComparison.Ordinal));
                if (pair?.OpenXmlPart is not ImagePart imagePart) {
                    result.Append(":missing");
                    continue;
                }
                result.Append(':').Append(imagePart.ContentType).Append(':');
                using Stream stream = imagePart.GetStream(FileMode.Open,
                    FileAccess.Read);
                using SHA256 sha = SHA256.Create();
                result.Append(Convert.ToBase64String(
                    sha.ComputeHash(stream)));
            }
            return result.ToString();
        }

        // SDK caller in LegacyPptShapeProjection.cs
        internal static string CreateTextFrameFingerprint(
            P.TextBody? textBody) => textBody?.BodyProperties?.OuterXml
                ?? string.Empty;
    }
}
