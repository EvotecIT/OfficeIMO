namespace OfficeIMO.Drawing;

internal sealed partial class OfficeAv1TransformReader {
    // Normative AOM v3.13.1 defaults represented as rising CDFs; every tile owns its mutable copies.
    private static int[][] CreateDepth() => new[] {
        new[] {19968,32768,0}, new[] {19968,32768,0}, new[] {24320,32768,0},
        new[] {12272,30172,32768,0}, new[] {12272,30172,32768,0}, new[] {18677,30848,32768,0},
        new[] {12986,15180,32768,0}, new[] {12986,15180,32768,0}, new[] {24302,25602,32768,0},
        new[] {5782,11475,32768,0}, new[] {5782,11475,32768,0}, new[] {16803,22759,32768,0}
    };
    private static int[][] CreateSplit() => new[] {
        new[] {28581,32768,0}, new[] {23846,32768,0}, new[] {20847,32768,0},
        new[] {24315,32768,0}, new[] {18196,32768,0}, new[] {12133,32768,0},
        new[] {18791,32768,0}, new[] {10887,32768,0}, new[] {11005,32768,0},
        new[] {27179,32768,0}, new[] {20004,32768,0}, new[] {11281,32768,0},
        new[] {26549,32768,0}, new[] {19308,32768,0}, new[] {14224,32768,0},
        new[] {28015,32768,0}, new[] {21546,32768,0}, new[] {14400,32768,0},
        new[] {28165,32768,0}, new[] {22401,32768,0}, new[] {16088,32768,0}
    };
}
