namespace OfficeIMO.DjVu;

public sealed partial class DjVuPage {
    internal Jb2Image DecodeMask(DjVuReadBudget budget) {
        IReadOnlyList<Jb2Bitmap>? dictionary = null;
        foreach (var chunk in DjVuDictionaryReader.Chain(Document, Component, budget.Cancellation))
            dictionary = new Jb2Decoder(chunk, budget, dictionary).Decode(true).Library;
        DjVuChunk? image = null;
        foreach (var chunk in Component.Form.Children) {
            if (chunk.Id == "Sjbz") {
                if (image != null) throw new InvalidDataException("Multiple JB2 page masks.");
                image = chunk;
            }
            if (chunk.Id == "Smmr") {
                if (image != null) throw new InvalidDataException("Multiple DjVu page masks.");
                image = chunk;
            }
        }
        if (image == null) throw new InvalidDataException("DjVu page has no JB2 mask.");
        var decoded = image.Id == "Smmr" ? DjVuMmrDecoder.Decode(image, budget) : new Jb2Decoder(image, budget, dictionary).Decode();
        if (decoded.Width != Width || decoded.Height != Height) throw new InvalidDataException("JB2 dimensions differ from page INFO.");
        return decoded;
    }
}
