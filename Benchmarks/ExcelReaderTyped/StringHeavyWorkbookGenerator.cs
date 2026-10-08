// Adapted upstream generator at ca5b50f99e8ef57ab476f0a2bc8043558d58b28d.
// Copyright (c) 2026 Gabriel Matte; MIT license in THIRD-PARTY-NOTICES.md.
using System.Globalization;
using System.Text;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Xls;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks
{
    internal static class StringHeavyWorkbookGenerator
    {
        internal static readonly string[] Headers = ["Code", "Category", "City", "Company", "PersonName", "Email", "Address", "Description", "Id", "Amount", "Date"];

        internal static object[] ExpectedValues(int index) {
            (string first, string last) = PersonName(index);
            return [Code(index), Categories[Spread(index, Categories.Length)], City(index), Company(index),
                $"{first} {last}", Email(first, last, index), Address(index), Description(index),
                (double)index, index * 2.75, RowDate(index)];
        }

        private const int CodePoolSize = 15_000;
        private const int CityPoolSize = 4_000;
        private const int CompanyPoolSize = 12_000;
        private const int NamePoolSize = 15_000;
        private const int DescriptionPoolSize = 20_000;

        private static readonly string[] Letters =
            ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M",
             "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];

        private static readonly string[] Consonants =
            ["b", "c", "d", "f", "g", "h", "j", "k", "l", "m", "n", "p", "r", "s", "t", "v", "w", "z",
             "th", "sh", "cr", "dr", "fl", "gr", "pl", "st", "tr"];

        private static readonly string[] Vowels = ["a", "e", "i", "o", "u", "ea", "io", "ou"];

        private static readonly string[] Categories =
            ["Electronics", "Furniture", "Apparel", "Groceries", "Automotive", "Hardware", "Software",
             "Toys", "Books", "Sporting Goods", "Office Supplies", "Health", "Beauty", "Garden",
             "Jewelry", "Footwear", "Pet Supplies", "Music", "Movies", "Appliances", "Tools",
             "Outdoor", "Baby", "Travel"];

        private static readonly string[] Suffixes = ["Inc", "LLC", "Group", "Holdings", "Partners", "Co"];

        private static readonly string[] Domains =
            ["example.com", "mail.com", "corp.net", "workmail.org", "bizmail.co", "teamhub.io",
             "cloudmail.net", "dataline.org", "swiftmail.com", "northgate.net", "bluepeak.org",
             "silverline.com", "oakridge.net", "brightpath.org", "fieldstone.com", "harborview.net",
             "summitworks.org", "clearwater.com", "irongate.net", "maplecrest.org", "stonebridge.com",
             "westfield.net", "eastgate.org", "truenorth.com"];

        private static readonly string[] States =
            ["AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA", "HI", "ID", "IL", "IN", "IA",
             "KS", "KY", "LA", "ME", "MD", "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ",
             "NM", "NY", "NC", "ND", "OH", "OK", "OR", "PA", "RI", "SC", "SD", "TN", "TX", "UT", "VT",
             "VA", "WA", "WV", "WI", "WY"];

        private static readonly string[] StreetTypes = ["St", "Ave", "Blvd", "Rd", "Ln", "Dr", "Ct", "Pl", "Way", "Ter"];

        public static Task<byte[]> BuildXlsxAsync(int rows)
        {
            return BuildAsync<XlsxWorkbookWriter, XlsxSheetWriter, XlsxRowWriter>(
                rows, static ms => XlsxWorkbookWriter.Create(ms, leaveOpen: true, options: new XlsxWriterOptions { UseSharedStrings = true }));
        }

        public static Task<byte[]> BuildXlsbAsync(int rows)
        {
            return BuildAsync<XlsbWorkbookWriter, XlsbSheetWriter, XlsbRowWriter>(
                rows, static ms => XlsbWorkbookWriter.Create(ms, leaveOpen: true, options: new XlsbWriterOptions { UseSharedStrings = true }));
        }

        public static async Task<byte[]> BuildXlsAsync(int rows)
        {
            await using MemoryStream ms = new();
            await using (XlsWorkbookWriter wb = XlsWorkbookWriter.Create(ms, leaveOpen: true))
            {
                XlsSheetWriter sheet = wb.AddSheet("S1");
                using (XlsRowWriter header = sheet.StartRow())
                {
                    WriteHeader(header);
                }
                for (int r = 1; r <= rows; r++)
                {
                    using XlsRowWriter row = sheet.StartRow();
                    WriteRow(row, r);
                }
                sheet.End();
                await wb.EndAsync();
            }
            return ms.ToArray();
        }

        public static byte[] BuildCsv(int rows)
        {
            var sb = new StringBuilder(rows * 160);
            sb.Append("Code,Category,City,Company,PersonName,Email,Address,Description,Id,Amount,Date\n");
            for (int r = 1; r <= rows; r++)
            {
                AppendCsvRow(sb, r);
            }
            return Encoding.UTF8.GetBytes(sb.ToString());
        }

        private static void AppendCsvRow(StringBuilder sb, int r)
        {
            sb.Append(Code(r)).Append(',')
              .Append(Categories[Spread(r, Categories.Length)]).Append(',')
              .Append(City(r)).Append(',')
              .Append(Company(r)).Append(',');
            (string first, string last) = PersonName(r);
            sb.Append(first).Append(' ').Append(last).Append(',')
              .Append(Email(first, last, r)).Append(',')
              .Append(Address(r)).Append(',')
              .Append(Description(r)).Append(',')
              .Append(r.ToString(CultureInfo.InvariantCulture)).Append(',')
              .Append((r * 2.75).ToString(CultureInfo.InvariantCulture)).Append(',')
              .Append(RowDate(r).ToString("O", CultureInfo.InvariantCulture)).Append('\n');
        }

        private static async Task<byte[]> BuildAsync<TWorkbook, TSheet, TRow>(
            int rows,
            Func<MemoryStream, TWorkbook> create)
            where TWorkbook : IWorkbookWriter<TSheet>
            where TSheet : ISheetWriter<TRow>
            where TRow : IRowWriter
        {
            await using var ms = new MemoryStream();
            await using (TWorkbook wb = create(ms))
            {
                TSheet sheet = wb.AddSheet("S1");
                await using (TRow header = await sheet.StartRowAsync())
                {
                    WriteHeader(header);
                }
                for (int r = 1; r <= rows; r++)
                {
                    await using TRow row = await sheet.StartRowAsync();
                    WriteRow(row, r);
                }
                await sheet.EndAsync();
                await wb.EndAsync();
            }
            return ms.ToArray();
        }

        private static void WriteHeader<TRow>(TRow row) where TRow : IRowWriter
        {
            row.Write("Code");
            row.Write("Category");
            row.Write("City");
            row.Write("Company");
            row.Write("PersonName");
            row.Write("Email");
            row.Write("Address");
            row.Write("Description");
            row.Write("Id");
            row.Write("Amount");
            row.Write("Date");
        }

        private static void WriteRow<TRow>(TRow row, int r) where TRow : IRowWriter
        {
            (string first, string last) = PersonName(r);
            row.Write(Code(r));
            row.Write(Categories[Spread(r, Categories.Length)]);
            row.Write(City(r));
            row.Write(Company(r));
            row.Write($"{first} {last}");
            row.Write(Email(first, last, r));
            row.Write(Address(r));
            row.Write(Description(r));
            row.Write(r);
            row.Write(r * 2.75);
            row.Write(RowDate(r));
        }

        private static DateTime RowDate(int r)
        {
            return DateTime.FromOADate(45292 + (r % 3650) + 0.5);
        }

        private static int Spread(int r, int modulus)
        {
            unchecked
            {
                uint h = (uint)r * 2654435761u;
                h ^= h >> 15;
                return (int)(h % (uint)modulus);
            }
        }

        private static uint Advance(uint state)
        {
            unchecked
            {
                return (state * 1664525u) + 1013904223u;
            }
        }

        private static string BuildWord(uint seed, int syllables)
        {
            uint state = seed;
            StringBuilder sb = new(syllables * 3);
            for (int i = 0; i < syllables; i++)
            {
                state = Advance(state);
                sb.Append(Consonants[state % Consonants.Length]);
                state = Advance(state);
                sb.Append(Vowels[state % Vowels.Length]);
            }
            sb[0] = char.ToUpperInvariant(sb[0]);
            return sb.ToString();
        }

        private static string Code(int r)
        {
            int idx = Spread(r, CodePoolSize);
            int digits = idx % 1000;
            int letterB = (idx / 1000) % 26;
            int letterC = idx / 26_000;
            return $"{Letters[letterC]}{Letters[letterB]}{digits:000}";
        }

        private static string City(int r)
        {
            int idx = Spread(r, CityPoolSize);
            string name = BuildWord((uint)idx * 2654435761u, 3);
            string state = States[Spread(idx, States.Length)];
            return $"{name} {state}";
        }

        private static string Company(int r)
        {
            int idx = Spread(r, CompanyPoolSize);
            string adjective = BuildWord((uint)idx * 2654435761u, 2);
            string noun = BuildWord((uint)idx * 40503u, 2);
            string suffix = Suffixes[Spread(idx, Suffixes.Length)];
            return $"{adjective} {noun} {suffix}";
        }

        private static (string First, string Last) PersonName(int r)
        {
            int idx = Spread(r, NamePoolSize);
            string first = BuildWord((uint)idx * 2654435761u, 2);
            string last = BuildWord((uint)idx * 97u, 2);
            return (first, last);
        }

        private static string Email(string first, string last, int r)
        {
            string domain = Domains[Spread((r * 7) + 3, Domains.Length)];
            char firstLower = char.ToLowerInvariant(first[0]);
            char lastLower = char.ToLowerInvariant(last[0]);
            return $"{firstLower}{first[1..]}.{lastLower}{last[1..]}@{domain}";
        }

        private static string Address(int r)
        {
            int number = 1 + (r % 9999);
            string street = BuildWord((uint)r * 2654435761u, 2);
            string type = StreetTypes[r % StreetTypes.Length];
            return $"{number} {street} {type}";
        }

        private static string Description(int r)
        {
            int idx = Spread(r, DescriptionPoolSize);
            int wordCount = 12 + (idx % 9);
            uint state = (uint)idx * 2654435761u + 1;
            StringBuilder sb = new(wordCount * 8);
            for (int i = 0; i < wordCount; i++)
            {
                state = Advance(state);
                if (i > 0) { sb.Append(' '); }
                sb.Append(BuildWord(state, 1 + (int)(state % 2)));
            }
            sb.Append('.');
            return sb.ToString();
        }
    }
}
