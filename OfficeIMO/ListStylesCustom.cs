using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO {
    public enum CustomListStyles {
        Bullet,
        Numbering
    }
    public static class CustomListStyle {
        public static NumberFormatValues GetStyle(CustomListStyles style) {
            switch (style) {
                case CustomListStyles.Bullet: return NumberFormatValues.Bullet;
                case CustomListStyles.Numbering: return NumberFormatValues.Decimal;
            }
            throw new ArgumentOutOfRangeException(nameof(style));
        }
    }
}
