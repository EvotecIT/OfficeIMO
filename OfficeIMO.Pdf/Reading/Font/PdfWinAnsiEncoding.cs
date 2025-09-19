namespace OfficeIMO.Pdf;

internal static class PdfWinAnsiEncoding {
    // Windows-1252 mapping (aka WinAnsiEncoding in PDF). Index is byte value.
    private static readonly char[] Map = new char[256] {
        '\u0000','\u0001','\u0002','\u0003','\u0004','\u0005','\u0006','\u0007','\b','\t','\n','\u000B','\f','\r','\u000E','\u000F',
        '\u0010','\u0011','\u0012','\u0013','\u0014','\u0015','\u0016','\u0017','\u0018','\u0019','\u001A','\u001B','\u001C','\u001D','\u001E','\u001F',
        ' ','!','"','#','$','%','&','\'','(',')','*','+',',','-','.','/',
        '0','1','2','3','4','5','6','7','8','9',':',';','<','=','>','?',
        '@','A','B','C','D','E','F','G','H','I','J','K','L','M','N','O',
        'P','Q','R','S','T','U','V','W','X','Y','Z','[','\\',']','^','_',
        '`','a','b','c','d','e','f','g','h','i','j','k','l','m','n','o',
        'p','q','r','s','t','u','v','w','x','y','z','{','|','}','~','\u007F',
        '€','\u0081','‚','ƒ','„','…','†','‡','ˆ','‰','Š','‹','Œ','\u008D','Ž','\u008F',
        '\u0090','‘','’','“','”','•','–','—','˜','™','š','›','œ','\u009D','ž','Ÿ',
        '\u00A0','¡','¢','£','¤','¥','¦','§','¨','©','ª','«','¬','\u00AD','®','¯',
        '°','±','²','³','´','µ','¶','·','¸','¹','º','»','¼','½','¾','¿',
        'À','Á','Â','Ã','Ä','Å','Æ','Ç','È','É','Ê','Ë','Ì','Í','Î','Ï',
        'Ð','Ñ','Ò','Ó','Ô','Õ','Ö','×','Ø','Ù','Ú','Û','Ü','Ý','Þ','ß',
        'à','á','â','ã','ä','å','æ','ç','è','é','ê','ë','ì','í','î','ï',
        'ð','ñ','ò','ó','ô','õ','ö','÷','ø','ù','ú','û','ü','ý','þ','ÿ'
    };

    private static readonly System.Collections.Generic.Dictionary<char, byte> ReverseMap = BuildReverse();

    public static string Decode(byte[] bytes) {
        var chars = new char[bytes.Length];
        for (int i = 0; i < bytes.Length; i++) chars[i] = Map[bytes[i]];
        return new string(chars);
    }

    public static byte[] Encode(string s) {
        var bytes = new byte[s.Length];
        for (int i = 0; i < s.Length; i++) {
            var ch = s[i];
            if (!ReverseMap.TryGetValue(ch, out var b)) b = (byte)'?';
            bytes[i] = b;
        }
        return bytes;
    }

    private static System.Collections.Generic.Dictionary<char, byte> BuildReverse() {
        var dict = new System.Collections.Generic.Dictionary<char, byte>();
        for (int i = 0; i < Map.Length; i++) {
            char c = Map[i];
            if (!dict.ContainsKey(c)) dict[c] = (byte)i;
        }
        return dict;
    }
}
