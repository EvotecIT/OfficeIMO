/* Independent component streams; use the pinned official AOM table and entropy implementation. */
#include <assert.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include "aom_dsp/entenc.h"
#include "aom_dsp/entdec.h"
#include "aom_dsp/prob.h"
#include "partition-defaults.h"

static int probability(const aom_cdf_prob *cdf, int symbol) { return cdf[symbol - 1] - cdf[symbol]; }

static aom_cdf_prob *select_cdf(aom_cdf_prob *cdf, aom_cdf_prob *edge, int n, int mode) {
    if (!mode) return cdf;
    int p = mode == 1 ? probability(cdf, 2) + probability(cdf, 3) + probability(cdf, 4) + probability(cdf, 6) + probability(cdf, 7)
                      : probability(cdf, 1) + probability(cdf, 3) + probability(cdf, 4) + probability(cdf, 5) + probability(cdf, 6);
    if (n == 10) p += probability(cdf, mode == 1 ? 9 : 8);
    edge[0] = (aom_cdf_prob)p; edge[1] = 0; edge[2] = 0;
    return edge;
}

static int symbol_at(int i, int n, int mode) { return mode ? (i / 3) & 1 : (i * 7 + i / 5) % n; }

static void make_case(FILE *out, int size_index, int context, int updates, int *first) {
    int n = size_index == 0 ? 4 : size_index == 4 ? 8 : 10;
    aom_cdf_prob enc_cdf[11], dec_cdf[11], edge[3];
    memcpy(enc_cdf, default_partition_cdf[size_index * 4 + context], sizeof(enc_cdf));
    memcpy(dec_cdf, enc_cdf, sizeof(enc_cdf));
    od_ec_enc enc; od_ec_enc_init(&enc, 128);
    const int operations = 48;
    for (int i = 0; i < operations; i++) {
        int mode = size_index ? i % 3 : 0, count = mode ? 2 : n;
        aom_cdf_prob *chosen = select_cdf(enc_cdf, edge, n, mode);
        int symbol = symbol_at(i, n, mode);
        od_ec_encode_cdf_q15(&enc, symbol, chosen, count);
        if (updates) update_cdf(chosen, symbol, count);
    }
    uint32_t size; unsigned char *bytes = od_ec_enc_done(&enc, &size); assert(bytes && !enc.error);
    od_ec_dec dec; od_ec_dec_init(&dec, bytes, size);
    for (int i = 0; i < operations; i++) {
        int mode = size_index ? i % 3 : 0, count = mode ? 2 : n;
        aom_cdf_prob *chosen = select_cdf(dec_cdf, edge, n, mode);
        int symbol = od_ec_decode_cdf_q15(&dec, chosen, count); assert(symbol == symbol_at(i, n, mode));
        if (updates) update_cdf(chosen, symbol, count);
    }
    assert(memcmp(enc_cdf, dec_cdf, sizeof(enc_cdf)) == 0);
    fprintf(out, "%s{\"pixels\":%d,\"context\":%d,\"updates\":%s,\"operations\":%d,\"hex\":\"",
            *first ? "" : ",", 8 << size_index, context, updates ? "true" : "false", operations); *first = 0;
    for (uint32_t i = 0; i < size; i++) fprintf(out, "%02x", bytes[i]);
    fprintf(out, "\"}"); od_ec_enc_clear(&enc);
}

static void prefix_case(char **argv) {
    FILE *input = fopen(argv[1], "rb"); assert(input);
    int offset = atoi(argv[2]), size = atoi(argv[3]); assert(offset >= 0 && size > 0);
    unsigned char *bytes = malloc((size_t)size); assert(bytes);
    assert(fseek(input, offset, SEEK_SET) == 0);
    assert(fread(bytes, 1, (size_t)size, input) == (size_t)size); fclose(input);
    int rows = atoi(argv[4]), cols = atoi(argv[5]), pixels = atoi(argv[6]);
    FILE *out = fopen(argv[7], "wb"); assert(out); fputc('[', out);
    od_ec_dec dec; od_ec_dec_init(&dec, bytes, (uint32_t)size);
    int first = 1;
    // Only the top-left descent is decoded. A non-split node reaches block syntax, which is outside this oracle.
    for (;;) {
        int half = pixels / 8, has_rows = half < rows, has_cols = half < cols, kind = 0;
        if (pixels >= 8) {
            if (!has_rows && !has_cols) kind = 3;
            else {
                int index = 0; for (int p = pixels; p > 8; p >>= 1) index++;
                int n = index == 0 ? 4 : index == 4 ? 8 : 10;
                aom_cdf_prob cdf[11], edge[3]; memcpy(cdf, default_partition_cdf[index * 4], sizeof(cdf));
                int mode = has_rows && has_cols ? 0 : has_cols ? 1 : 2;
                int symbol = od_ec_decode_cdf_q15(&dec, select_cdf(cdf, edge, n, mode), mode ? 2 : n);
                kind = mode ? symbol ? 3 : mode == 1 ? 1 : 2 : symbol;
            }
        }
        fprintf(out, "%s{\"pixels\":%d,\"kind\":%d}", first ? "" : ",", pixels, kind); first = 0;
        if (kind != 3) break;
        pixels >>= 1;
    }
    fprintf(out, "]\n"); fclose(out); free(bytes);
}

int main(int argc, char **argv) {
    if (argc == 8) { prefix_case(argv); return 0; }
    assert(argc == 2);
    FILE *out = fopen(argv[1], "wb"); assert(out);
    fprintf(out, "{\"producer\":\"AOM v3.13.1 official default partition tables and entropy encoder/decoder\",\"nativeSelfCheck\":true,\"cases\":[");
    int first = 1;
    for (int size = 0; size < 5; size++) for (int context = 0; context < 4; context++)
        for (int updates = 0; updates <= 1; updates++) make_case(out, size, context, updates, &first);
    fprintf(out, "]}\n"); fclose(out); return 0;
}
