using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        ///     Gets or sets whether the slide is hidden in slide show mode.
        /// </summary>
        public bool Hidden {
            get {
                if (SlideRoot.Show?.Value != null) {
                    return !SlideRoot.Show.Value;
                }

                return IsHiddenShowValue(GetLegacySlideIdShowValue(GetSlideId()));
            }
            set {
                SlideId slideId = GetSlideId();
                slideId.RemoveAttribute("show", string.Empty);

                if (value) {
                    SlideRoot.Show = false;
                } else {
                    SlideRoot.Show = null;
                }
            }
        }

        /// <summary>
        ///     Hides the slide in slide show mode.
        /// </summary>
        public void Hide() => Hidden = true;

        /// <summary>
        ///     Shows the slide in slide show mode.
        /// </summary>
        public void Show() => Hidden = false;

        private SlideId GetSlideId() {
            PresentationPart presentationPart = _slidePart.GetParentParts()
                .OfType<PresentationPart>()
                .FirstOrDefault()
                ?? throw new InvalidOperationException("Slide is not attached to a presentation.");

            SlideIdList? slideIdList = presentationPart.Presentation?.SlideIdList;
            if (slideIdList == null) {
                throw new InvalidOperationException("Presentation has no slide list.");
            }

            string relId = presentationPart.GetIdOfPart(_slidePart);
            SlideId? slideId = slideIdList.Elements<SlideId>()
                .FirstOrDefault(id => id.RelationshipId?.Value == relId);

            if (slideId == null) {
                throw new InvalidOperationException("Slide not found in presentation.");
            }

            return slideId;
        }

        private static string? GetLegacySlideIdShowValue(SlideId slideId) {
            return slideId.GetAttributes()
                .FirstOrDefault(attribute =>
                    attribute.LocalName == "show" && string.IsNullOrEmpty(attribute.NamespaceUri))
                .Value;
        }

        private static bool IsHiddenShowValue(string? showValue) {
            if (string.IsNullOrEmpty(showValue)) {
                return false;
            }

            return string.Equals(showValue, "0", StringComparison.Ordinal) ||
                   string.Equals(showValue, "false", StringComparison.OrdinalIgnoreCase);
        }

        private void NormalizeHiddenSlideMarkup() {
            SlideId slideId = GetSlideId();
            string? legacyShowValue = GetLegacySlideIdShowValue(slideId);
            if (SlideRoot.Show?.Value == null && IsHiddenShowValue(legacyShowValue)) {
                SlideRoot.Show = false;
            }

            slideId.RemoveAttribute("show", string.Empty);
        }
    }
}
