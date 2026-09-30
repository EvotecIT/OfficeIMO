/* Independent test-data generator. Compile only against opt-in AOM v3.13.1 entropy sources. */
#include <assert.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include "aom_dsp/entenc.h"
#include "aom_dsp/entdec.h"
#include "aom_dsp/prob.h"

static void initial_cdf(aom_cdf_prob *cdf, int n, int degenerate) {
    const int four[] = { 19132, 25510, 30392, 32768 };
    const int ten[] = { 15597, 20929, 24571, 26706, 27664, 28821, 29601, 30571, 31902, 32768 };
    for (int i = 0; i < n; i++) {
        int v = n == 4 ? four[i] : n == 10 ? ten[i] : (i + 1) * 32768 / n;
        if (degenerate) v = i < n - 1 ? 100 : 32768;
        cdf[i] = 32768 - v;
    }
    cdf[n] = 0;
}

static void print_cdf(FILE *out, aom_cdf_prob *cdf, int n) {
    fputc('[', out);
    for (int i = 0; i <= n; i++) fprintf(out, "%s%d", i ? "," : "", i == n ? cdf[i] : 32768 - cdf[i]);
    fputc(']', out);
}

static int symbol_at(int i, int n) { return (i * 13 + i / 7) % n; }
static int literal_at(int i) { return (i * 139 + 71) & 0x7fffffff; }

static void make_case(FILE *out, int n, int updates, int degenerate, int operations, int *first) {
    aom_cdf_prob start[17] = {0}, enc_cdf[17], dec_cdf[17];
    initial_cdf(start, n, degenerate);
    memcpy(enc_cdf, start, sizeof(start)); memcpy(dec_cdf, start, sizeof(start));
    od_ec_enc enc; od_ec_enc_init(&enc, 1024);
    for (int i = 0; i < operations; i++) {
        int symbol = symbol_at(i, n);
        od_ec_encode_cdf_q15(&enc, symbol, enc_cdf, n);
        if (updates) update_cdf(enc_cdf, symbol, n);
        if (i % 5 == 0) od_ec_encode_bool_q15(&enc, (i / 5) & 1, 16384);
        if (i % 32 == 0) for (int b = 30; b >= 0; b--) od_ec_encode_bool_q15(&enc, (literal_at(i) >> b) & 1, 16384);
    }
    uint32_t size; unsigned char *bytes = od_ec_enc_done(&enc, &size);
    assert(bytes && !enc.error);
    od_ec_dec dec; od_ec_dec_init(&dec, bytes, size);
    for (int i = 0; i < operations; i++) {
        int symbol = od_ec_decode_cdf_q15(&dec, dec_cdf, n);
        assert(symbol == symbol_at(i, n));
        if (updates) update_cdf(dec_cdf, symbol, n);
        if (i % 5 == 0) assert(od_ec_decode_bool_q15(&dec, 16384) == ((i / 5) & 1));
        if (i % 32 == 0) {
            int value = 0;
            for (int b = 0; b < 31; b++) value = (value << 1) | od_ec_decode_bool_q15(&dec, 16384);
            assert(value == literal_at(i));
        }
    }
    assert(memcmp(enc_cdf, dec_cdf, sizeof(aom_cdf_prob) * (n + 1)) == 0);
    fprintf(out, "%s{\"symbols\":%d,\"updates\":%s,\"degenerate\":%s,\"operations\":%d,\"initialCdf\":", *first ? "" : ",", n, updates ? "true" : "false", degenerate ? "true" : "false", operations);
    *first = 0; print_cdf(out, start, n);
    fprintf(out, ",\"finalCdf\":"); print_cdf(out, enc_cdf, n);
    fprintf(out, ",\"hex\":\""); for (uint32_t i = 0; i < size; i++) fprintf(out, "%02x", bytes[i]);
    fprintf(out, "\",\"nativeTellBits\":%d}", od_ec_dec_tell(&dec));
    od_ec_enc_clear(&enc);
}

static void boolean_case(FILE *out, const char *bits, int *first) {
    od_ec_enc enc; od_ec_enc_init(&enc, 128);
    for (const char *p = bits; *p; p++) od_ec_encode_bool_q15(&enc, *p == '1', 16384);
    uint32_t size; unsigned char *bytes = od_ec_enc_done(&enc, &size); assert(bytes && !enc.error);
    od_ec_dec dec; od_ec_dec_init(&dec, bytes, size);
    for (const char *p = bits; *p; p++) assert(od_ec_decode_bool_q15(&dec, 16384) == (*p == '1'));
    fprintf(out, "%s{\"bits\":\"%s\",\"hex\":\"", *first ? "" : ",", bits); *first = 0;
    for (uint32_t i = 0; i < size; i++) fprintf(out, "%02x", bytes[i]);
    fprintf(out, "\"}"); od_ec_enc_clear(&enc);
}

static void prefix_case(char **argv) {
    FILE *input = fopen(argv[1], "rb"); assert(input);
    int offset = atoi(argv[2]), size = atoi(argv[3]); assert(offset >= 0 && size > 0);
    unsigned char *bytes = malloc((size_t)size); assert(bytes);
    assert(fseek(input, offset, SEEK_SET) == 0);
    assert(fread(bytes, 1, (size_t)size, input) == (size_t)size); fclose(input);
    const int initial[] = { 20137, 21547, 23078, 29566, 29837, 30261, 30524, 30892, 31724, 32768, 0 };
    aom_cdf_prob cdf[11];
    for (int i = 0; i < 10; i++) cdf[i] = 32768 - initial[i]; cdf[10] = 0;
    od_ec_dec dec; od_ec_dec_init(&dec, bytes, (uint32_t)size);
    int symbol = od_ec_decode_cdf_q15(&dec, cdf, 10); update_cdf(cdf, symbol, 10);
    FILE *out = fopen(argv[4], "wb"); assert(out);
    fprintf(out, "{\"symbol\":%d,\"finalCdf\":", symbol); print_cdf(out, cdf, 10); fprintf(out, "}\n");
    fclose(out); free(bytes);
}

int main(int argc, char **argv) {
    if (argc == 5) { prefix_case(argv); return 0; }
    assert(argc == 2);
    FILE *out = fopen(argv[1], "wb"); assert(out);
    fprintf(out, "{\"producer\":\"AOM v3.13.1 od_ec_encode_cdf_q15/od_ec_encode_bool_q15\",\"nativeSelfCheck\":true,\"cases\":[");
    int first = 1; const int counts[] = { 2, 3, 4, 10, 16 };
    for (int i = 0; i < 5; i++) for (int updates = 0; updates <= 1; updates++) make_case(out, counts[i], updates, 0, 256, &first);
    make_case(out, 4, 0, 1, 64, &first);
    make_case(out, 2, 0, 0, 1, &first);
    fprintf(out, "],\"booleanCases\":["); first = 1;
    boolean_case(out, "0", &first); boolean_case(out, "1", &first); boolean_case(out, "00101011", &first);
    boolean_case(out, "1111111111111111", &first);
    fprintf(out, "]}\n"); fclose(out);
    return 0;
}
