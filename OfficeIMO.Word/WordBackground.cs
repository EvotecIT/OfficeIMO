using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public class WordBackground {


        public WordBackground() {

            DocumentBackground documentBackground1 = new DocumentBackground() { Color = "B671F5" };

            DocumentBackground documentBackground2 = new DocumentBackground() { Color = "BF8F00", ThemeColor = ThemeColorValues.Accent4, ThemeShade = "BF" };


        }
    }
}
