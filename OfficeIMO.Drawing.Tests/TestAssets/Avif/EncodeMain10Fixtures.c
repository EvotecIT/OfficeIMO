/* Opt-in libavif 1.4.2 producer: AOM encoding and independent dav1d decoding.
 * The driver fills native 8/10-bit planes directly; it contains no codec arithmetic. */
#include "avif/avif.h"
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

static void require(int condition, const char * message) {
    if (!condition) { fprintf(stderr, "%s\n", message); exit(1); }
}

static FILE * openOutput(const char * folder, const char * name, const char * extension) {
    char path[4096];
    int count = snprintf(path, sizeof(path), "%s/%s.%s", folder, name, extension);
    require(count > 0 && count < (int)sizeof(path), "Output path exceeds driver limit");
    FILE * file = fopen(path, "wb");
    require(file != NULL, "Cannot create fixture output");
    return file;
}

static void writePlane(FILE * file, const uint8_t * plane, uint32_t pitch, int width, int height, int depth) {
    for (int y = 0; y < height; y++) {
        const uint16_t * row = (const uint16_t *)(plane + y * pitch);
        for (int x = 0; x < width; x++) {
            unsigned int value = depth > 8 ? row[x] : plane[y * pitch + x];
            require(value < (1u << depth), "Decoded sample exceeds source range");
            require(fputc(value & 255, file) != EOF && fputc(value >> 8, file) != EOF, "Cannot write plane");
        }
    }
}

static void writePlanes(const char * folder, const char * name, const char * extension, const avifImage * image) {
    FILE * file = openOutput(folder, name, extension);
    int width = (int)image->width, height = (int)image->height;
    writePlane(file, image->yuvPlanes[0], image->yuvRowBytes[0], width, height, image->depth);
    if (image->yuvFormat == AVIF_PIXEL_FORMAT_YUV420)
        for (int p = 1; p < 3; p++)
            writePlane(file, image->yuvPlanes[p], image->yuvRowBytes[p], (width + 1) / 2, (height + 1) / 2, image->depth);
    if (image->alphaPlane) writePlane(file, image->alphaPlane, image->alphaRowBytes, width, height, image->depth);
    require(fclose(file) == 0, "Cannot close plane output");
}

int main(int argc, char ** argv) {
    require((argc == 2 || argc == 4) && strcmp(avifVersion(), "1.4.2") == 0, "Use libavif 1.4.2, output folder, optional depth and name suffix");
    int depth = argc == 4 ? atoi(argv[2]) : 10;
    const char * suffix = argc == 4 ? argv[3] : "";
    require(depth == 8 || depth == 10, "Only Main eight/ten-bit fixtures are supported");
    require(avifCodecName(AVIF_CODEC_CHOICE_AOM, AVIF_CODEC_FLAG_CAN_ENCODE) != NULL, "AOM encoder unavailable");
    require(avifCodecName(AVIF_CODEC_CHOICE_DAV1D, AVIF_CODEC_FLAG_CAN_DECODE) != NULL, "dav1d decoder unavailable");
    char versions[256]; avifCodecVersions(versions);
    printf("{\"libavif\":\"%s\",\"codecs\":\"%s\",\"cases\":[", avifVersion(), versions);
    int first = 1, width = 49, height = 33;
    for (int mono = 0; mono < 2; mono++) for (int full = 0; full < 2; full++) for (int alpha = 0; alpha < 2; alpha++) {
        char name[128];
        snprintf(name, sizeof(name), "avif-main%d%s-%s-%s%s", depth, suffix, mono ? "mono" : "420", full ? "full" : "limited", alpha ? "-alpha" : "");
        avifImage * image = avifImageCreate(width, height, depth, mono ? AVIF_PIXEL_FORMAT_YUV400 : AVIF_PIXEL_FORMAT_YUV420);
        require(image != NULL, "Cannot allocate source image");
        image->colorPrimaries = AVIF_COLOR_PRIMARIES_BT709;
        image->transferCharacteristics = AVIF_TRANSFER_CHARACTERISTICS_SRGB;
        image->matrixCoefficients = AVIF_MATRIX_COEFFICIENTS_BT601;
        image->yuvRange = full ? AVIF_RANGE_FULL : AVIF_RANGE_LIMITED;
        require(avifImageAllocatePlanes(image, alpha ? AVIF_PLANES_ALL : AVIF_PLANES_YUV) == AVIF_RESULT_OK, "Cannot allocate source planes");
        for (int p = 0; p < (mono ? 1 : 3); p++) {
            int w = p ? (width + 1) / 2 : width, h = p ? (height + 1) / 2 : height;
            for (int y = 0; y < h; y++) {
                uint16_t * row = (uint16_t *)(image->yuvPlanes[p] + y * image->yuvRowBytes[p]);
                for (int x = 0; x < w; x++) {
                    int value = (x * 67 + y * 113 + p * 193 + (x / 7) * 157) % 1024;
                    int scale = 1 << (depth - 8);
                    int sample = full ? value * ((1 << depth) - 1) / 1023 : 16 * scale + value * (p ? 224 : 219) * scale / 1023;
                    if (depth > 8) row[x] = (uint16_t)sample; else image->yuvPlanes[p][y * image->yuvRowBytes[p] + x] = (uint8_t)sample;
                }
            }
        }
        if (alpha) for (int y = 0; y < height; y++) {
            uint16_t * row = (uint16_t *)(image->alphaPlane + y * image->alphaRowBytes);
            for (int x = 0; x < width; x++) {
                int sample = ((x * 53 + y * 29) % 1024) * ((1 << depth) - 1) / 1023;
                if (depth > 8) row[x] = (uint16_t)sample; else image->alphaPlane[y * image->alphaRowBytes + x] = (uint8_t)sample;
            }
        }
        writePlanes(argv[1], name, "source-yuv16", image);
        avifEncoder * encoder = avifEncoderCreate(); require(encoder != NULL, "Cannot allocate encoder");
        encoder->codecChoice = AVIF_CODEC_CHOICE_AOM; encoder->maxThreads = 1; encoder->speed = 6;
        encoder->quality = 85; encoder->qualityAlpha = 100;
        avifRWData encoded = AVIF_DATA_EMPTY;
        avifResult encodedResult = avifEncoderWrite(encoder, image, &encoded);
        if (encodedResult != AVIF_RESULT_OK)
            fprintf(stderr, "%s: %s; %s\n", name, avifResultToString(encodedResult), encoder->diag.error);
        require(encodedResult == AVIF_RESULT_OK, "Cannot encode native still fixture");
        FILE * file = openOutput(argv[1], name, "avif");
        require(fwrite(encoded.data, 1, encoded.size, file) == encoded.size && fclose(file) == 0, "Cannot write AVIF");
        avifDecoder * decoder = avifDecoderCreate(); require(decoder != NULL, "Cannot allocate decoder");
        decoder->codecChoice = AVIF_CODEC_CHOICE_DAV1D; decoder->maxThreads = 1;
        avifImage * decoded = avifImageCreateEmpty(); require(decoded != NULL, "Cannot allocate decoded image");
        require(avifDecoderReadMemory(decoder, decoded, encoded.data, encoded.size) == AVIF_RESULT_OK, "Independent dav1d decode failed");
        require(decoded->width == (uint32_t)width && decoded->height == (uint32_t)height && decoded->depth == (uint32_t)depth &&
            decoded->yuvFormat == image->yuvFormat && decoded->yuvRange == image->yuvRange &&
            (decoded->alphaPlane != NULL) == alpha, "Decoded profile differs from requested fixture");
        writePlanes(argv[1], name, "yuv16", decoded);
        avifRGBImage rgb; avifRGBImageSetDefaults(&rgb, decoded);
        rgb.depth = 8; rgb.format = AVIF_RGB_FORMAT_RGBA;
        rgb.avoidLibYUV = AVIF_TRUE; rgb.chromaUpsampling = AVIF_CHROMA_UPSAMPLING_BILINEAR;
        require(avifRGBImageAllocatePixels(&rgb) == AVIF_RESULT_OK && avifImageYUVToRGB(decoded, &rgb) == AVIF_RESULT_OK,
            "Independent RGBA conversion failed");
        file = openOutput(argv[1], name, "rgba");
        for (int y = 0; y < height; y++) require(fwrite(rgb.pixels + y * rgb.rowBytes, 1, width * 4, file) == (size_t)(width * 4), "Cannot write RGBA");
        require(fclose(file) == 0, "Cannot close RGBA output");
        printf("%s{\"name\":\"%s\",\"width\":%d,\"height\":%d,\"bitDepth\":%d,\"monochrome\":%s,\"fullRange\":%s,\"hasAlpha\":%s}",
            first ? "" : ",", name, width, height, depth, mono ? "true" : "false", full ? "true" : "false", alpha ? "true" : "false");
        first = 0;
        avifRGBImageFreePixels(&rgb); avifImageDestroy(decoded); avifDecoderDestroy(decoder);
        avifRWDataFree(&encoded); avifEncoderDestroy(encoder); avifImageDestroy(image);
    }
    printf("]}\n"); return 0;
}
