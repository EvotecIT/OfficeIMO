namespace OfficeIMO.Invoicing;

// Exact dictionaries from FA3.xsd (B646B6B5...) and KodyKrajow_v10-0E.xsd (1D41A1B3...).
// These are schema codes, including historical/special entries, not a current ISO registry.
internal static class Fa3ScalarContract {
    internal static readonly HashSet<string> Currencies = new HashSet<string>((
        "AED AFN ALL AMD ANG AOA ARS AUD AWG AZN BAM BBD BDT BGN BHD BIF BMD BND BOB BOV BRL BSD BTN BWP BYN BZD CAD CDF CHE CHF CHW CLF CLP CNY COP COU CRC CUC CUP CVE CZK DJF DKK DOP DZD EGP ERN ETB EUR FJD FKP GBP GEL GGP GHS GIP GMD GNF GTQ GYD HKD HNL HRK HTG HUF IDR ILS IMP INR IQD IRR ISK JEP JMD JOD JPY KES KGS KHR KMF KPW KRW KWD KYD KZT LAK LBP LKR LRD LSL LYD MAD MDL MGA MKD MMK MNT MOP MRU MUR MVR MWK MXN MXV MYR MZN NAD NGN NIO NOK NPR NZD OMR PAB PEN PGK PHP PKR PLN PYG QAR RON RSD RUB RWF SAR SBD SCR SDG SEK SGD SHP SLL SOS SRD SSP STN SVC SYP SZL THB TJS TMT TND TOP TRY TTD TWD TZS UAH UGX USD USN UYI UYU UYW UZS VES VND VUV WST XAF XAG XAU XBA XBB XBC XBD XCD XCG XDR XOF XPD XPF XPT XSU XUA XXX YER ZAR ZMW ZWL"
        ).Split(' '), StringComparer.Ordinal);
    internal static readonly HashSet<string> Countries = new HashSet<string>((
        "AF AX AL DZ AD AO AI AQ AG AN SA AR AM AW AU AT AZ BS BH BD BB BE BZ BJ BM BT BY BO BQ BA BW BR BN IO BG BF BI XC CL CN HR CW CY TD ME DK DM DO DJ EG EC ER EE ET FK FJ PH FI FR TF GA GM GH GI GR GD GL GE GU GG GY GF GP GT GN GQ GW HT ES HN HK IN ID IQ IR IE IS IL JM JP YE JE JO KY KH CM CA QA KZ KE KG KI CO KM CG CD KP XK CR CU KW LA LS LB LR LY LI LT LV LU MK MG YT MO MW MV MY ML MT MP MA MQ MR MU MX XL FM UM MD MC MN MS MZ MM NA NR NP NL DE NE NG NI NU NF NO NC NZ PS OM PK PW PA PG PY PE PN PF PL GS PT PR CF CZ KR ZA RE RU RO RW EH BL KN LC MF VC SV WS AS SM SN RS SC SL SG SK SI SO LK PM US SZ SD SS SR SJ SH SY CH SE TJ TH TW TZ TG TK TO TT TN TR TM TV UG UA UY UZ VU WF VA HU VE GB VN IT TL CI BV CX IM SX CK VI VG HM CC MH FO SB ST TC ZM CV ZW AE XI"
        ).Split(' '), StringComparer.Ordinal);

    internal static void Date(string path, DateTime? value, int firstYear, Action<string, string> error) {
        if (value.HasValue && (value.Value.Date < new DateTime(firstYear, firstYear == 2016 ? 7 : 1, 1) || value.Value.Date > new DateTime(2050, 1, 1)))
            error(path, "Date is outside the pinned schema range.");
    }
}
