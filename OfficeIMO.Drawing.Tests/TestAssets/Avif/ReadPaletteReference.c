/* Binary case transport only; the observation entry point uses AOM's decoder. */
#include <stdint.h>
#include <stdio.h>
#include <stdlib.h>

void *office_palette_create(const uint8_t *, int, int, int, int, int, int, int);
void office_palette_leaf(void *, const int *, FILE *);
void office_palette_destroy(void *);

static int read_int(FILE *in) {
  uint8_t b[4];
  if (fread(b, 1, 4, in) != 4) abort();
  return (int)(b[0] | (uint32_t)b[1] << 8 | (uint32_t)b[2] << 16 |
               (uint32_t)b[3] << 24);
}
int main(int argc, char **argv) {
  if (argc != 2) abort();
  FILE *in = fopen(argv[1], "rb");
  if (!in) abort();
  int cases = read_int(in);
  for (int c = 0; c < cases; ++c) {
    int depth = read_int(in), extent = read_int(in), updates = read_int(in);
    int screen = read_int(in), filter = read_int(in), mono = read_int(in);
    int length = read_int(in), states = read_int(in);
    if (length <= 0 || length > 1000000 || states < 1 || states > 1024) abort();
    uint8_t *bytes = malloc(length);
    if (!bytes || fread(bytes, 1, length, in) != (size_t)length) abort();
    void *p = office_palette_create(bytes, length, depth, extent, updates,
                                   screen, filter, mono);
    fputs("[", stdout);
    for (int s = 0; s < states; ++s) {
      int v[13];
      for (int i = 0; i < 13; ++i) v[i] = read_int(in);
      if (s) fputc(',', stdout);
      office_palette_leaf(p, v, stdout);
    }
    office_palette_destroy(p);
    free(bytes);
    puts("]");
  }
  if (fgetc(in) != EOF) abort();
  fclose(in);
  return 0;
}
