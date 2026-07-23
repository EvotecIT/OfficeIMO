namespace OfficeIMO.Pdf;

internal static class PdfMacRomanEncoding {
    private static readonly char[] Map = new char[256] {
        '\u0000','\u0001','\u0002','\u0003','\u0004','\u0005','\u0006','\u0007','\b','\t','\n','\u000B','\f','\r','\u000E','\u000F',
        '\u0010','\u0011','\u0012','\u0013','\u0014','\u0015','\u0016','\u0017','\u0018','\u0019','\u001A','\u001B','\u001C','\u001D','\u001E','\u001F',
        ' ','!','"','#','$','%','&','\'','(',')','*','+',',','-','.','/',
        '0','1','2','3','4','5','6','7','8','9',':',';','<','=','>','?',
        '@','A','B','C','D','E','F','G','H','I','J','K','L','M','N','O',
        'P','Q','R','S','T','U','V','W','X','Y','Z','[','\\',']','^','_',
        '`','a','b','c','d','e','f','g','h','i','j','k','l','m','n','o',
        'p','q','r','s','t','u','v','w','x','y','z','{','|','}','~','\u007F',
        '\u00C4','\u00C5','\u00C7','\u00C9','\u00D1','\u00D6','\u00DC','\u00E1','\u00E0','\u00E2','\u00E4','\u00E3','\u00E5','\u00E7','\u00E9','\u00E8',
        '\u00EA','\u00EB','\u00ED','\u00EC','\u00EE','\u00EF','\u00F1','\u00F3','\u00F2','\u00F4','\u00F6','\u00F5','\u00FA','\u00F9','\u00FB','\u00FC',
        '\u2020','\u00B0','\u00A2','\u00A3','\u00A7','\u2022','\u00B6','\u00DF','\u00AE','\u00A9','\u2122','\u00B4','\u00A8','\u2260','\u00C6','\u00D8',
        '\u221E','\u00B1','\u2264','\u2265','\u00A5','\u00B5','\u2202','\u2211','\u220F','\u03C0','\u222B','\u00AA','\u00BA','\u03A9','\u00E6','\u00F8',
        '\u00BF','\u00A1','\u00AC','\u221A','\u0192','\u2248','\u2206','\u00AB','\u00BB','\u2026','\u00A0','\u00C0','\u00C3','\u00D5','\u0152','\u0153',
        '\u2013','\u2014','\u201C','\u201D','\u2018','\u2019','\u00F7','\u25CA','\u00FF','\u0178','\u2044','\u20AC','\u2039','\u203A','\uFB01','\uFB02',
        '\u2021','\u00B7','\u201A','\u201E','\u2030','\u00C2','\u00CA','\u00C1','\u00CB','\u00C8','\u00CD','\u00CE','\u00CF','\u00CC','\u00D3','\u00D4',
        '\uF8FF','\u00D2','\u00DA','\u00DB','\u00D9','\u0131','\u02C6','\u02DC','\u00AF','\u02D8','\u02D9','\u02DA','\u00B8','\u02DD','\u02DB','\u02C7'
    };

    public static string Decode(byte[] bytes) {
        return Decode(bytes, int.MaxValue);
    }

    public static string Decode(byte[] bytes, int maxOutputCharacters) {
        if (bytes.LongLength > maxOutputCharacters) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.DecodedTextCharacters, maxOutputCharacters, bytes.LongLength);
        }
        var chars = new char[bytes.Length];
        for (int i = 0; i < bytes.Length; i++) chars[i] = Map[bytes[i]];
        return new string(chars);
    }
}
