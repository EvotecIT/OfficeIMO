using System.Globalization;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Tests.Pdf;

/// <summary>One painting operator on the first page: its kind, device-independent placement and fill color.</summary>
internal sealed record PdfViewerPaint(string Kind, double A, double B, double C, double D, double E, double F, string Color);

/// <summary>
/// Follows the page content of the first page the way viewers do, independently of OfficeIMO's renderer: the
/// concatenated streams start with an empty save stack and a Q with nothing saved is ignored, so content that restores
/// more states than it saved keeps whatever it changed afterwards.
/// </summary>
internal static class PdfViewerGeometry {
    internal static IReadOnlyList<PdfViewerPaint> Paint(byte[] pdf) {
        Dictionary<int, PdfIndirectObject> objects = PdfSyntax.ParseObjects(pdf, null).Map;
        int pageNumber = PdfReadDocument.Open(pdf).Pages[0].ObjectNumber;
        var page = (PdfDictionary)objects[pageNumber].Value;
        var content = new MemoryStream();
        PdfObject? contents = page.Items.TryGetValue("Contents", out PdfObject? value) ? PdfObjectLookup.Resolve(objects, value) : null;
        IEnumerable<PdfObject> entries = contents is PdfArray array ? array.Items : contents is null ? Array.Empty<PdfObject>() : new[] { value! };
        foreach (PdfObject entry in entries) {
            byte[] decoded = StreamDecoder.Decode((PdfStream)PdfObjectLookup.Resolve(objects, entry)!, objects);
            content.Write(decoded, 0, decoded.Length);
            content.WriteByte((byte)'\n');
        }

        var paints = new List<PdfViewerPaint>();
        // Fill color plus the state a paint inherits without naming it: extended state, text rendering and scaling, clips.
        var saved = new Stack<(Matrix Ctm, string Fill, string Inherited)>();
        Matrix ctm = Matrix.Identity, textMatrix = Matrix.Identity, lineMatrix = Matrix.Identity;
        string fill = "0 g", inherited = string.Empty;
        bool clipPending = false;
        var path = new List<(double X, double Y)>();
        PdfContentStreamInterpreter.Interpret(PdfContentSource.FromBytes(content.ToArray()), int.MaxValue, operation => {
            double N(int index) => Convert.ToDouble(operation.Operands[index], CultureInfo.InvariantCulture);
            if (operation.Name is "f" or "F" or "f*" or "B" or "B*" or "b" or "b*" or "S" or "s" or "n" && clipPending) {
                inherited += " clip[" + string.Join(" ", path.Select(p => Fmt(Round(p.X)) + "," + Fmt(Round(p.Y)))) + "]";
                clipPending = false;
            }
            switch (operation.Name) {
                case "q": saved.Push((ctm, fill, inherited)); break;
                case "Q": if (saved.Count > 0) (ctm, fill, inherited) = saved.Pop(); break;
                case "cm": ctm = new Matrix(N(0), N(1), N(2), N(3), N(4), N(5)).Then(ctm); break;
                case "gs": inherited += " gs " + operation.Operands[0]; break;
                case "Tr": inherited += " " + Fmt(N(0)) + " Tr"; break;
                case "Tz": inherited += " " + Fmt(N(0)) + " Tz"; break;
                case "W": case "W*": clipPending = true; break;
                case "rg": fill = Fmt(N(0)) + " " + Fmt(N(1)) + " " + Fmt(N(2)) + " rg"; break;
                case "g": fill = Fmt(N(0)) + " g"; break;
                case "re":
                    double x = N(0), y = N(1), w = N(2), h = N(3);
                    foreach ((double px, double py) in new[] { (x, y), (x + w, y), (x, y + h), (x + w, y + h) }) path.Add(ctm.Transform(px, py));
                    break;
                case "m": case "l": path.Add(ctm.Transform(N(0), N(1))); break;
                case "c": path.Add(ctm.Transform(N(4), N(5))); break;
                case "f": case "F": case "f*": case "B": case "B*": case "b": case "b*": case "S": case "s":
                    if (path.Count > 0) {
                        paints.Add(new PdfViewerPaint("path", Round(path.Min(p => p.X)), Round(path.Min(p => p.Y)),
                            Round(path.Max(p => p.X)), Round(path.Max(p => p.Y)), 0, 0, fill + inherited));
                    }
                    path.Clear();
                    break;
                case "n": path.Clear(); break;
                case "BT": textMatrix = lineMatrix = Matrix.Identity; break;
                case "Tm": textMatrix = lineMatrix = new Matrix(N(0), N(1), N(2), N(3), N(4), N(5)); break;
                case "Td": case "TD": textMatrix = lineMatrix = new Matrix(1, 0, 0, 1, N(0), N(1)).Then(lineMatrix); break;
                case "Tj": case "TJ": case "'": case "\"":
                    paints.Add(Placed("text", textMatrix.Then(ctm), fill + inherited));
                    break;
                case "Do": paints.Add(Placed("xobject", ctm, fill + inherited)); break;
                case "BI": paints.Add(Placed("inline-image", ctm, fill + inherited)); break;
            }
        });
        return paints;
    }

    private static PdfViewerPaint Placed(string kind, Matrix m, string fill) =>
        new(kind, Round(m.A), Round(m.B), Round(m.C), Round(m.D), Round(m.E), Round(m.F), fill);

    private static double Round(double value) => Math.Round(value, 2) + 0D;

    private static string Fmt(double value) => value.ToString("0.###", CultureInfo.InvariantCulture);

    private readonly struct Matrix {
        internal Matrix(double a, double b, double c, double d, double e, double f) { A = a; B = b; C = c; D = d; E = e; F = f; }
        internal static Matrix Identity => new(1, 0, 0, 1, 0, 0);
        internal double A { get; }
        internal double B { get; }
        internal double C { get; }
        internal double D { get; }
        internal double E { get; }
        internal double F { get; }

        /// <summary>Applies this matrix first, then <paramref name="next"/>.</summary>
        internal Matrix Then(Matrix next) => new(
            A * next.A + B * next.C, A * next.B + B * next.D,
            C * next.A + D * next.C, C * next.B + D * next.D,
            E * next.A + F * next.C + next.E, E * next.B + F * next.D + next.F);

        internal (double X, double Y) Transform(double x, double y) => (x * A + y * C + E, x * B + y * D + F);
    }
}
