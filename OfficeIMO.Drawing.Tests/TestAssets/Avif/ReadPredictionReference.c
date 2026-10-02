/* Portable opt-in observation driver. Pixel algorithms stay in the pinned native library. */
#include <stdint.h>
#include <stdio.h>
#include <stdlib.h>
#include "av1/common/common_data.h"
void office_intra_prediction(const int *, const uint8_t *, const uint8_t *, uint8_t *);
void office_cfl_prediction(const int *, const uint8_t *, uint8_t *);
void office_palette_prediction(const int *, const uint8_t *, const uint8_t *, uint8_t *);
void office_intra_prediction_high(const int *, const uint16_t *, const uint16_t *, uint16_t *, int);
void office_cfl_prediction_high(const int *, const uint16_t *, uint16_t *, int);
void office_palette_prediction_high(const int *, const uint16_t *, const uint8_t *, uint16_t *);
static int integer(FILE *f) {
  uint8_t b[4]; if (fread(b, 1, 4, f) != 4) exit(2);
  return (int32_t)((uint32_t)b[0] | (uint32_t)b[1] << 8 | (uint32_t)b[2] << 16 | (uint32_t)b[3] << 24);
}
static void bytes(FILE *f, uint8_t *dst, int count, int capacity) {
  if (count < 0 || count > capacity || fread(dst, 1, (size_t)count, f) != (size_t)count) exit(3);
}
static void samples(FILE *f, uint16_t *dst, int count, int capacity, int bit_depth) {
  if (count < 0 || count > capacity) exit(3);
  for (int i = 0; i < count; i++) {
    int lo = fgetc(f), hi = bit_depth == 8 ? 0 : fgetc(f);
    if (lo == EOF || hi == EOF) exit(3);
    dst[i] = (uint16_t)(lo | hi << 8);
    if (dst[i] >= (1 << bit_depth)) exit(3);
  }
}
int main(int argc, char **argv) {
  if (argc != 3) return 1;
  int bit_depth = atoi(argv[2]); if (bit_depth != 8 && bit_depth != 10) return 1;
  FILE *f = fopen(argv[1], "rb"); if (!f) return 2;
  int count = integer(f);
  for (int i = 0; i < count; i++) {
    int kind = integer(f), p[16]; uint8_t a[4096], b[4096], output[4096];
    uint16_t ah[4096], bh[4096], oh[4096];
    for (int j = 0; j < 16; j++) p[j] = integer(f);
    if (p[0] < 0 || p[0] >= TX_SIZES_ALL) return 4;
    if (bit_depth == 10) {
      if (kind == 0) {samples(f, ah, p[11], 128, bit_depth); samples(f, bh, p[12], 128, bit_depth); office_intra_prediction_high(p, ah, bh, oh, bit_depth);}
      else if (kind == 1) {samples(f, ah, p[8], 4096, bit_depth); office_cfl_prediction_high(p, ah, oh, bit_depth);}
      else if (kind == 2) {samples(f, ah, p[6], 8, bit_depth); bytes(f, b, p[2] * p[3], 4096); office_palette_prediction_high(p, ah, b, oh);}
      else return 5;
    }
    else if (kind == 0) {bytes(f, a, p[11], 128); bytes(f, b, p[12], 128); office_intra_prediction(p, a, b, output);}
    else if (kind == 1) {bytes(f, a, p[8], 4096); office_cfl_prediction(p, a, output);}
    else if (kind == 2) {bytes(f, a, p[6], 8); bytes(f, b, p[2] * p[3], 4096); office_palette_prediction(p, a, b, output);}
    else return 5;
    int pixels = tx_size_wide[p[0]] * tx_size_high[p[0]];
    printf("{\"scenario\":%d,\"pixels\":[", i);
    for (int j = 0; j < pixels; j++) printf("%s%d", j ? "," : "", bit_depth == 8 ? output[j] : oh[j]);
    puts("]}");
  }
  if (fgetc(f) != EOF) return 6;
  fclose(f); return 0;
}
