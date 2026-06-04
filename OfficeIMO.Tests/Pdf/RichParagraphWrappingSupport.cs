using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf {
    public partial class RichParagraphWrappingTests {

        private static object InvokeWrapRichRuns(IEnumerable<TextRun> runs, double maxWidthPts, double fontSize, PdfStandardFont baseFont, double? tabStopWidth = null) {
            var method = typeof(PdfWriter).GetMethod("WrapRichRuns", BindingFlags.NonPublic | BindingFlags.Static);
            Assert.NotNull(method);
            return method!.Invoke(null, new object?[] { runs, maxWidthPts, fontSize, baseFont, fontSize * 1.4, null, tabStopWidth ?? 36.0 })!;
        }

        private static T InvokePrivateFontMethod<T>(string methodName, params object[] parameters) {
            var method = typeof(PdfWriter).GetMethod(methodName, BindingFlags.NonPublic | BindingFlags.Static);
            Assert.NotNull(method);
            return (T)method!.Invoke(null, parameters)!;
        }

        private static TargetInvocationException InvokePrivateFontMethodExpectingFailure(string methodName, params object[] parameters) {
            var method = typeof(PdfWriter).GetMethod(methodName, BindingFlags.NonPublic | BindingFlags.Static);
            Assert.NotNull(method);
            return Assert.Throws<TargetInvocationException>(() => method!.Invoke(null, parameters));
        }

        private static List<List<object>> ExtractLines(object wrapResult) {
            var item1Field = wrapResult.GetType().GetField("Item1");
            Assert.NotNull(item1Field);
            var item1 = item1Field!.GetValue(wrapResult)!;
            var lines = new List<List<object>>();
            foreach (var lineObj in (IEnumerable)item1) {
                var segs = new List<object>();
                foreach (var segObj in (IEnumerable)lineObj) segs.Add(segObj);
                lines.Add(segs);
            }
            return lines;
        }

        private static PdfStandardFont ExtractFont(object seg) {
            var prop = seg.GetType().GetProperty("Font");
            Assert.NotNull(prop);
            return (PdfStandardFont)prop!.GetValue(seg)!;
        }

        private static string ExtractText(object seg) {
            var prop = seg.GetType().GetProperty("Text");
            Assert.NotNull(prop);
            return (string)prop!.GetValue(seg)!;
        }

        private static bool ExtractBold(object seg) {
            var prop = seg.GetType().GetProperty("Bold");
            Assert.NotNull(prop);
            return (bool)prop!.GetValue(seg)!;
        }

        private static bool ExtractLeadingSpace(object seg) {
            var prop = seg.GetType().GetProperty("LeadingSpace");
            Assert.NotNull(prop);
            return (bool)prop!.GetValue(seg)!;
        }

        private static double ExtractLeadingAdvance(object seg) {
            var prop = seg.GetType().GetProperty("LeadingAdvance");
            Assert.NotNull(prop);
            return (double)prop!.GetValue(seg)!;
        }

        private static bool ExtractLeadingSpaceIsExpandable(object seg) {
            var prop = seg.GetType().GetProperty("LeadingSpaceIsExpandable");
            Assert.NotNull(prop);
            return (bool)prop!.GetValue(seg)!;
        }

        private static PdfTabLeaderStyle ExtractLeadingTabLeader(object seg) {
            var prop = seg.GetType().GetProperty("LeadingTabLeader");
            Assert.NotNull(prop);
            return (PdfTabLeaderStyle)prop!.GetValue(seg)!;
        }

        private static PdfTextBaseline ExtractBaseline(object seg) {
            var prop = seg.GetType().GetProperty("Baseline");
            Assert.NotNull(prop);
            return (PdfTextBaseline)prop!.GetValue(seg)!;
        }
    }
}
