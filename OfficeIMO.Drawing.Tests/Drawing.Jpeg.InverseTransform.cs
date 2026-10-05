using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingJpegInverseTransformTests {
    // Native libjpeg grayscale fixtures exercise dense coefficients at three
    // quantization levels in baseline and progressive scans. Expected samples
    // preserve the decoder's integer rounding across hardware implementations.
    [Theory]
    [MemberData(nameof(GrayscaleFixtures))]
    public void IntegerTransformPreservesQualifiedPixelSamples(string name, string encoded, string expectedSamples) {
        var actual = OfficeJpegCodec.Decode(Convert.FromBase64String(encoded));
        Assert.Equal(8, actual.Width);
        Assert.Equal(8, actual.Height);
        byte[] samples = Convert.FromBase64String(expectedSamples);
        var expected = new byte[64 * 4];
        for (int index = 0; index < samples.Length; index++) {
            expected[index * 4] = samples[index];
            expected[index * 4 + 1] = samples[index];
            expected[index * 4 + 2] = samples[index];
            expected[index * 4 + 3] = 255;
        }
        Assert.Equal(expected, actual.GetPixels());
    }

    public static IEnumerable<object[]> GrayscaleFixtures() {
        yield return new object[] { "idct-gray-q1-baseline",
            "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAP//////////////////////////////////////////////////////////////////////////////////////wAALCAAIAAgBAREA/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/9oACAEBAAA/AE9v8/8A6uenev/Z",
            "DTzsQYF2xwBjIEh1/wgAlGQcdoCwN1tJL2a/RgaW3ABRmkE6emokkZxlCoBQSiW3nWE4iwB4oWxzxBQ/AIo5sw==" };
        yield return new object[] { "idct-gray-q85-baseline",
            "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAUDBAQEAwUEBAQFBQUGBwwIBwcHBw8LCwkMEQ8SEhEPERETFhwXExQaFRERGCEYGh0dHx8fExciJCIeJBweHx7/wAALCAAIAAgBAREA/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/9oACAEBAAA/AMXTNJj0fSbHw6Vms7RbFfPQxvHFeTSRIXM0kZ3SW6x3EKiAYaWSQptb5XT/2Q==",
            "ASdFbpK84QITQ3Cr2QAjZyRYlMwAVpK4LIm8CFiF0x0+jeswgd0of1ytEHG+K4DYZ9IwmQBn0DR242DKQ7YgmA==" };
        yield return new object[] { "idct-gray-q100-baseline",
            "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/wAALCAAIAAgBAREA/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/9oACAEBAAA/APn3wP8ACfTvg98Pfhx+zu9r4g8BeCLP4P6Oni7SLrQ/FXg/wX8bfiF8R/hn4N1rxnffGv4g/DzxBd+KfiP+y/4F+FH7RfwZ8F6V+y/4YtfDXxN/ad+P3xR1D4Vad4L8ZWx8D/FP4Sf/2Q==",
            "ACVKb5S53gMRQnGh0QExYSNdmNQOSYS/M3m/BEuR1x5EleY3iNkqe1WxDWnFIXzZZs40mwNp0Dd36VvNP7EjlQ==" };
        yield return new object[] { "idct-gray-q1-progressive",
            "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAP//////////////////////////////////////////////////////////////////////////////////////wgALCAAIAAgBAREA/8QAFAABAAAAAAAAAAAAAAAAAAAAAf/aAAgBAQAAAAE//8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQABBQJ//8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQAGPwJ//8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQABPyF//9oACAEBAAAAEP8A/8QAGhAAAAcAAAAAAAAAAAAAAAAAACExQWHw8f/aAAgBAQABPxCLho4//9k=",
            "DTzsQYF2xwBjIEh1/wgAlGQcdoCwN1tJL2a/RgaW3ABRmkE6emokkZxlCoBQSiW3nWE4iwB4oWxzxBQ/AIo5sw==" };
        yield return new object[] { "idct-gray-q85-progressive",
            "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAUDBAQEAwUEBAQFBQUGBwwIBwcHBw8LCwkMEQ8SEhEPERETFhwXExQaFRERGCEYGh0dHx8fExciJCIeJBweHx7/wgALCAAIAAgBAREA/8QAFAABAAAAAAAAAAAAAAAAAAAABP/aAAgBAQAAAAEX/8QAFRABAQAAAAAAAAAAAAAAAAAAAwH/2gAIAQEAAQUCMoJf/8QAHBAAAgICAwAAAAAAAAAAAAAAAQIRIQBBBBIx/9oACAEBAAY/Ak49ovS6gOSNkerDCtk5/8QAGBABAQADAAAAAAAAAAAAAAAAAREhMUH/2gAIAQEAAT8hmcpmgS7gB7gjhP/aAAgBAQAAABB//8QAFhABAQEAAAAAAAAAAAAAAAAAAREA/9oACAEBAAE/EGBRGqNogAYpo//Z",
            "ASdFbpK84QITQ3Cr2QAjZyRYlMwAVpK4LIm8CFiF0x0+jeswgd0of1ytEHG+K4DYZ9IwmQBn0DR242DKQ7YgmA==" };
        yield return new object[] { "idct-gray-q100-progressive",
            "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/wgALCAAIAAgBAREA/8QAFAABAAAAAAAAAAAAAAAAAAAAB//aAAgBAQAAAAE+/8QAFhABAQEAAAAAAAAAAAAAAAAABQME/9oACAEBAAEFAsJMxz//xAAaEAACAwEBAAAAAAAAAAAAAAAEBQIDBhQS/9oACAEBAAY/Al2d8kABQT09dUqCgwnbBisDuMm6YLyJFMcuCq0SYKrLjRGZ6d+0sVVhGR4Win//xAAVEAEBAAAAAAAAAAAAAAAAAAABAP/aAAgBAQABPyFDEi6N9V4xBP8A/9oACAEBAAAAEH//xAAVEAEBAAAAAAAAAAAAAAAAAAAAAf/aAAgBAQABPxCmpigou5ldZP8A/9k=",
            "ACVKb5S53gMRQnGh0QExYSNdmNQOSYS/M3m/BEuR1x5EleY3iNkqe1WxDWnFIXzZZs40mwNp0Dd36VvNP7EjlQ==" };
    }
}
