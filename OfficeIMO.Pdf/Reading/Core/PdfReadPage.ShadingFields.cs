using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private const int MaximumMeshVertices = 2_000_000;
    private const int MaximumMeshTriangles = 4_000_000;
    private const int MaximumPatchSubdivision = 16;
    private static readonly double[] MeshParameterDomain = { 0D, 1D };

    /// <summary>
    /// Reads axial and radial shadings whose Extend, BBox or circle geometry the drawing gradients cannot represent,
    /// and every mesh shading, as exact colour fields.
    /// </summary>
    private bool TryReadShadingField(PdfDictionary dictionary, PdfStream? stream, int shadingType, PdfPageColorSpace colorSpace,
        IReadOnlyList<double>? coords, double[] shadingDomain, bool extendStart, bool extendEnd,
        OfficeIccRenderingIntent intent, PageContentBudget budget, out PdfPageShadingResource shading) {
        shading = default;
        double[]? bounds = null;
        if (dictionary.Items.TryGetValue("BBox", out PdfObject? boundsObject) && ResolveObject(boundsObject) is not PdfNull) {
            if (!TryReadExactFiniteNumberArray(boundsObject, 4, out double[] parsed)) return false;
            bounds = new[] { Math.Min(parsed[0], parsed[2]), Math.Min(parsed[1], parsed[3]), Math.Max(parsed[0], parsed[2]), Math.Max(parsed[1], parsed[3]) };
        }
        double[]? background = null;
        if (dictionary.Items.TryGetValue("Background", out PdfObject? backgroundObject) && ResolveObject(backgroundObject) is not PdfNull) {
            if (!TryReadExactFiniteNumberArray(backgroundObject, colorSpace.ComponentCount, out double[] parsedBackground)) return false;
            background = parsedBackground;
        }
        PdfObject? functionObject = dictionary.Items.TryGetValue("Function", out PdfObject? function) ? function : null;
        PdfShadingColorFunctions? functions = null;
        if (functionObject != null && ResolveObject(functionObject) is not PdfNull) {
            if (!TryReadShadingColorFunctions(functionObject, colorSpace.ComponentCount, shadingDomain, budget, out functions)) return false;
        } else if (shadingType is 2 or 3) {
            return false;
        }
        IPdfShadingFieldSampler sampler;
        if (shadingType == 2) {
            sampler = new PdfAxialShadingSampler(coords![0], coords[1], coords[2], coords[3], extendStart, extendEnd, functions!);
        } else if (shadingType == 3) {
            if (coords![2] < 0D || coords[5] < 0D) return false;
            sampler = new PdfRadialShadingSampler(coords[0], coords[1], coords[2], coords[3], coords[4], coords[5], extendStart, extendEnd, functions!);
        } else {
            if (stream == null || !TryReadMeshSampler(dictionary, stream, shadingType, colorSpace, functions, budget, out PdfMeshShadingSampler? mesh)) return false;
            sampler = mesh!;
        }
        shading = new PdfPageShadingResource(new PdfFunctionShading(sampler, colorSpace, bounds, background, intent,
            EffectiveOutputIntentColorTransform, budget.ChargeFunctionShadingWork, budget.CancellationToken,
            budget.FunctionShadingScale, budget.ChargeFunctionShadingPixels, dictionary), supportsExactType3Projection: false);
        return true;
    }

    private bool TryReadShadingColorFunctions(PdfObject functionObject, int componentCount, double[] domain, PageContentBudget budget,
        out PdfShadingColorFunctions? functions) {
        functions = null;
        PdfArray? array = ResolveArray(functionObject);
        if (array != null && array.Items.Count != componentCount) return false;
        var resolved = new PdfColorFunction[array == null ? 1 : array.Items.Count];
        for (int index = 0; index < resolved.Length; index++) {
            if (!PdfColorSpaceFunctionResolver.TryCreateFunction(array == null ? functionObject : array.Items[index], 1,
                    array == null ? componentCount : 1, _objects, _limits.MaxDecodedStreamBytes,
                    budget.ColorFunctionResolutionContext, out PdfColorFunction parsed)) return false;
            resolved[index] = parsed;
        }
        functions = new PdfShadingColorFunctions(resolved, domain[0], domain[1]);
        return true;
    }

    // ISO 32000-1, 8.7.4.5.5 to 8.7.4.5.8: free-form and lattice Gouraud triangles, Coons and tensor-product patches.
    private bool TryReadMeshSampler(PdfDictionary dictionary, PdfStream stream, int shadingType, PdfPageColorSpace colorSpace,
        PdfShadingColorFunctions? functions, PageContentBudget budget, out PdfMeshShadingSampler? sampler) {
        sampler = null;
        int bitsPerCoordinate = TryReadInteger(dictionary.Items.TryGetValue("BitsPerCoordinate", out PdfObject? bpc) ? bpc : null) ?? 0;
        int bitsPerComponent = TryReadInteger(dictionary.Items.TryGetValue("BitsPerComponent", out PdfObject? bpcomp) ? bpcomp : null) ?? 0;
        int bitsPerFlag = shadingType == 5 ? 0 : TryReadInteger(dictionary.Items.TryGetValue("BitsPerFlag", out PdfObject? bpf) ? bpf : null) ?? 0;
        if (bitsPerCoordinate is not (1 or 2 or 4 or 8 or 12 or 16 or 24 or 32) ||
            bitsPerComponent is not (1 or 2 or 4 or 8 or 12 or 16) ||
            shadingType != 5 && bitsPerFlag is not (2 or 4 or 8)) return false;
        int valueCount = functions != null ? 1 : colorSpace.ComponentCount;
        if (!dictionary.Items.TryGetValue("Decode", out PdfObject? decodeObject) ||
            !TryReadExactFiniteNumberArray(decodeObject, 4 + 2 * valueCount, out double[] decode)) return false;
        int verticesPerRow = 0;
        if (shadingType == 5) {
            verticesPerRow = TryReadInteger(dictionary.Items.TryGetValue("VerticesPerRow", out PdfObject? vpr) ? vpr : null) ?? 0;
            if (verticesPerRow < 2) return false;
        }
        if (Filters.StreamDecoder.GetUnsupportedFilters(stream.Dictionary, _objects).Count != 0 ||
            !Filters.StreamDecoder.TryDecode(stream, _limits.MaxDecodedStreamBytes, out byte[] data, _objects)) return false;

        var reader = new MeshBitReader(data);
        var xs = new List<double>();
        var ys = new List<double>();
        var values = new List<double>();
        var triangles = new List<int>();
        var vertexValues = new double[valueCount];
        bool ReadVertex(int flagBits, out int flag) {
            flag = 0;
            if (flagBits > 0) {
                if (!reader.TryRead(flagBits, out uint rawFlag)) return false;
                flag = (int)rawFlag;
            }
            if (!reader.TryReadScaled(bitsPerCoordinate, decode[0], decode[1], out double x) ||
                !reader.TryReadScaled(bitsPerCoordinate, decode[2], decode[3], out double y)) return false;
            for (int component = 0; component < valueCount; component++) {
                if (!reader.TryReadScaled(bitsPerComponent, decode[4 + 2 * component], decode[5 + 2 * component], out vertexValues[component])) return false;
            }
            if (xs.Count >= MaximumMeshVertices) throw new InvalidDataException("The mesh shading has too many vertices.");
            xs.Add(x); ys.Add(y);
            for (int component = 0; component < valueCount; component++) values.Add(vertexValues[component]);
            return true;
        }
        void AddTriangle(int a, int b, int c) {
            if (triangles.Count / 3 >= MaximumMeshTriangles) throw new InvalidDataException("The mesh shading has too many triangles.");
            triangles.Add(a); triangles.Add(b); triangles.Add(c);
        }

        if (shadingType == 4) {
            int va = -1, vb = -1, vc = -1, pending = 0;
            while (!reader.AtEnd) {
                budget.CancellationToken.ThrowIfCancellationRequested();
                if (!ReadVertex(bitsPerFlag, out int flag)) break;
                reader.Align();
                int vertex = xs.Count - 1;
                if (pending > 0 || flag == 0 && vc < 0) {
                    if (pending == 0) pending = 3;
                    if (pending == 3) va = vertex; else if (pending == 2) vb = vertex; else vc = vertex;
                    if (--pending == 0) AddTriangle(va, vb, vc);
                } else if (flag == 0) {
                    va = vertex;
                    pending = 2;
                } else if (va < 0 || vb < 0 || vc < 0) {
                    return false;
                } else if (flag == 1) {
                    va = vb; vb = vc; vc = vertex;
                    AddTriangle(va, vb, vc);
                } else if (flag == 2) {
                    vb = vc; vc = vertex;
                    AddTriangle(va, vb, vc);
                }
            }
        } else if (shadingType == 5) {
            while (!reader.AtEnd) {
                budget.CancellationToken.ThrowIfCancellationRequested();
                if (!ReadVertex(0, out _)) break;
                reader.Align();
            }
            int rows = xs.Count / verticesPerRow;
            for (int row = 0; row + 1 < rows; row++) {
                budget.CancellationToken.ThrowIfCancellationRequested();
                for (int column = 0; column + 1 < verticesPerRow; column++) {
                    if ((column & 1023) == 0) budget.CancellationToken.ThrowIfCancellationRequested();
                    int a = row * verticesPerRow + column, b = a + 1, c = a + verticesPerRow, d = c + 1;
                    AddTriangle(a, b, c);
                    AddTriangle(b, d, c);
                }
            }
        } else {
            ReadPatches(reader, shadingType == 7, bitsPerFlag, bitsPerCoordinate, bitsPerComponent, decode, valueCount,
                xs, ys, values, triangles, budget);
        }
        if (triangles.Count == 0) return false;
        sampler = new PdfMeshShadingSampler(xs, ys, values, valueCount, triangles, functions, budget.CancellationToken);
        return true;
    }

    private static void ReadPatches(MeshBitReader reader, bool tensor, int bitsPerFlag, int bitsPerCoordinate, int bitsPerComponent,
        double[] decode, int valueCount, List<double> xs, List<double> ys, List<double> values, List<int> triangles, PageContentBudget budget) {
        // Points p[i, j] with i along u and j along v; colours at the corners (0,0), (0,3), (3,3), (3,0).
        var px = new double[4, 4];
        var py = new double[4, 4];
        var colors = new double[4][];
        double[,]? previousX = null, previousY = null;
        double[][]? previousColors = null;
        int[] boundaryI = { 0, 0, 0, 0, 1, 2, 3, 3, 3, 3, 2, 1 };
        int[] boundaryJ = { 0, 1, 2, 3, 3, 3, 3, 2, 1, 0, 0, 0 };
        int[] interiorI = { 1, 1, 2, 2 };
        int[] interiorJ = { 1, 2, 2, 1 };
        var patches = new List<(double[,] X, double[,] Y, double[][] Colors)>();
        while (!reader.AtEnd) {
            budget.CancellationToken.ThrowIfCancellationRequested();
            if (!reader.TryRead(bitsPerFlag, out uint rawFlag)) break;
            int flag = (int)rawFlag;
            if (flag > 3 || flag != 0 && previousX == null) break;
            int firstBoundary = 0;
            int firstColor = 0;
            if (flag != 0) {
                // The shared edge of the previous patch becomes points 1 to 4 and colours 1 and 2 of this one.
                int[] edge = flag switch { 1 => new[] { 3, 4, 5, 6 }, 2 => new[] { 6, 7, 8, 9 }, _ => new[] { 9, 10, 11, 0 } };
                for (int index = 0; index < 4; index++) {
                    px[boundaryI[index], boundaryJ[index]] = previousX![boundaryI[edge[index]], boundaryJ[edge[index]]];
                    py[boundaryI[index], boundaryJ[index]] = previousY![boundaryI[edge[index]], boundaryJ[edge[index]]];
                }
                int[] colorEdge = flag switch { 1 => new[] { 1, 2 }, 2 => new[] { 2, 3 }, _ => new[] { 3, 0 } };
                colors[0] = (double[])previousColors![colorEdge[0]].Clone();
                colors[1] = (double[])previousColors[colorEdge[1]].Clone();
                firstBoundary = 4;
                firstColor = 2;
            }
            bool complete = true;
            for (int index = firstBoundary; index < 12 && complete; index++) {
                complete = reader.TryReadScaled(bitsPerCoordinate, decode[0], decode[1], out px[boundaryI[index], boundaryJ[index]]) &&
                    reader.TryReadScaled(bitsPerCoordinate, decode[2], decode[3], out py[boundaryI[index], boundaryJ[index]]);
            }
            if (tensor) {
                for (int index = 0; index < 4 && complete; index++) {
                    complete = reader.TryReadScaled(bitsPerCoordinate, decode[0], decode[1], out px[interiorI[index], interiorJ[index]]) &&
                        reader.TryReadScaled(bitsPerCoordinate, decode[2], decode[3], out py[interiorI[index], interiorJ[index]]);
                }
            }
            for (int color = firstColor; color < 4 && complete; color++) {
                colors[color] = new double[valueCount];
                for (int component = 0; component < valueCount && complete; component++) {
                    complete = reader.TryReadScaled(bitsPerComponent, decode[4 + 2 * component], decode[5 + 2 * component], out colors[color][component]);
                }
            }
            reader.Align();
            if (!complete) break;
            if (!tensor) {
                // Coons interior control points from the boundary (ISO 32000-1, 8.7.4.5.8).
                px[1, 1] = CoonsInterior(px, 0, 0, 0, 1, 1, 0, 0, 3, 3, 0, 3, 1, 1, 3, 3, 3);
                py[1, 1] = CoonsInterior(py, 0, 0, 0, 1, 1, 0, 0, 3, 3, 0, 3, 1, 1, 3, 3, 3);
                px[1, 2] = CoonsInterior(px, 0, 3, 0, 2, 1, 3, 0, 0, 3, 3, 3, 2, 1, 0, 3, 0);
                py[1, 2] = CoonsInterior(py, 0, 3, 0, 2, 1, 3, 0, 0, 3, 3, 3, 2, 1, 0, 3, 0);
                px[2, 1] = CoonsInterior(px, 3, 0, 3, 1, 2, 0, 3, 3, 0, 0, 0, 1, 2, 3, 0, 3);
                py[2, 1] = CoonsInterior(py, 3, 0, 3, 1, 2, 0, 3, 3, 0, 0, 0, 1, 2, 3, 0, 3);
                px[2, 2] = CoonsInterior(px, 3, 3, 3, 2, 2, 3, 3, 0, 0, 3, 0, 2, 2, 0, 0, 0);
                py[2, 2] = CoonsInterior(py, 3, 3, 3, 2, 2, 3, 3, 0, 0, 3, 0, 2, 2, 0, 0, 0);
            }
            // Every patch needs at least a 3x3 vertex grid and eight triangles. Bound retention before cloning.
            if (patches.Count >= Math.Min(MaximumMeshVertices / 9, MaximumMeshTriangles / 8))
                throw new InvalidDataException("The mesh shading has too many patches.");
            patches.Add(((double[,])px.Clone(), (double[,])py.Clone(), new[] { colors[0], colors[1], colors[2], colors[3] }));
            previousX = (double[,])px.Clone();
            previousY = (double[,])py.Clone();
            previousColors = new[] { colors[0], colors[1], colors[2], colors[3] };
        }
        if (patches.Count == 0) return;
        int subdivision = Math.Max(2, Math.Min(MaximumPatchSubdivision, Math.Min(
            (int)Math.Sqrt(MaximumMeshTriangles / (2D * patches.Count)), (int)Math.Sqrt(MaximumMeshVertices / (double)patches.Count) - 1)));
        var bu = new double[4];
        var bv = new double[4];
        foreach (var patch in patches) {
            budget.CancellationToken.ThrowIfCancellationRequested();
            int baseVertex = xs.Count;
            for (int row = 0; row <= subdivision; row++) {
                budget.CancellationToken.ThrowIfCancellationRequested();
                double v = row / (double)subdivision;
                Bernstein(v, bv);
                for (int column = 0; column <= subdivision; column++) {
                    double u = column / (double)subdivision;
                    Bernstein(u, bu);
                    double x = 0D, y = 0D;
                    for (int i = 0; i < 4; i++) for (int j = 0; j < 4; j++) {
                        x += bu[i] * bv[j] * patch.X[i, j];
                        y += bu[i] * bv[j] * patch.Y[i, j];
                    }
                    if (xs.Count >= MaximumMeshVertices) throw new InvalidDataException("The mesh shading has too many vertices.");
                    xs.Add(x); ys.Add(y);
                    for (int component = 0; component < valueCount; component++) {
                        values.Add((1D - u) * (1D - v) * patch.Colors[0][component] + (1D - u) * v * patch.Colors[1][component] +
                            u * v * patch.Colors[2][component] + u * (1D - v) * patch.Colors[3][component]);
                    }
                }
            }
            for (int row = 0; row < subdivision; row++) {
                budget.CancellationToken.ThrowIfCancellationRequested();
                for (int column = 0; column < subdivision; column++) {
                    int a = baseVertex + row * (subdivision + 1) + column, b = a + 1, c = a + subdivision + 1, d = c + 1;
                    triangles.Add(a); triangles.Add(b); triangles.Add(c);
                    triangles.Add(b); triangles.Add(d); triangles.Add(c);
                }
            }
        }
    }

    // (-4 p_a + 6 (p_b + p_c) - 2 (p_d + p_e) + 3 (p_f + p_g) - p_h) / 9
    private static double CoonsInterior(double[,] p, int ai, int aj, int bi, int bj, int ci, int cj, int di, int dj, int ei, int ej,
        int fi, int fj, int gi, int gj, int hi, int hj) =>
        (-4D * p[ai, aj] + 6D * (p[bi, bj] + p[ci, cj]) - 2D * (p[di, dj] + p[ei, ej]) + 3D * (p[fi, fj] + p[gi, gj]) - p[hi, hj]) / 9D;

    private static void Bernstein(double t, double[] result) {
        double s = 1D - t;
        result[0] = s * s * s;
        result[1] = 3D * t * s * s;
        result[2] = 3D * t * t * s;
        result[3] = t * t * t;
    }

    private sealed class MeshBitReader {
        private readonly byte[] _data;
        private long _bit;

        internal MeshBitReader(byte[] data) => _data = data;

        internal bool AtEnd => _bit >= (long)_data.Length * 8L;

        internal void Align() => _bit = (_bit + 7L) & ~7L;

        internal bool TryRead(int bits, out uint value) {
            value = 0;
            if (_bit + bits > (long)_data.Length * 8L) return false;
            ulong accumulated = 0;
            for (int index = 0; index < bits; index++) {
                long position = _bit + index;
                int bit = (_data[position >> 3] >> (7 - (int)(position & 7))) & 1;
                accumulated = (accumulated << 1) | (uint)bit;
            }
            _bit += bits;
            value = (uint)accumulated;
            return true;
        }

        internal bool TryReadScaled(int bits, double minimum, double maximum, out double value) {
            value = 0D;
            if (!TryRead(bits, out uint raw)) return false;
            double maximumRaw = bits == 32 ? uint.MaxValue : (1UL << bits) - 1UL;
            value = minimum + raw * (maximum - minimum) / maximumRaw;
            return true;
        }
    }
}
