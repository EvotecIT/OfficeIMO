// Opt-in independent eight-bit AVIF chroma reference; no codec or conversion arithmetic.
// Compile with libavif 1.4.2 and the pinned libyuv revision in the evidence receipt.
#include "avif/avif.h"
#include "libyuv.h"
#include <cstdio>
#include <cstring>
#include <vector>

int main(int argc, char** argv) {
    if (argc != 3 || std::strcmp(avifVersion(), "1.4.2") != 0) return 1;
    avifDecoder* decoder = avifDecoderCreate();
    avifImage* image = avifImageCreateEmpty();
    if (!decoder || !image) return 2;
    decoder->codecChoice = AVIF_CODEC_CHOICE_DAV1D;
    decoder->maxThreads = 1;
    if (avifDecoderReadFile(decoder, image, argv[1]) != AVIF_RESULT_OK ||
        image->depth != 8 || image->yuvFormat != AVIF_PIXEL_FORMAT_YUV420 ||
        image->yuvRange != AVIF_RANGE_FULL ||
        image->matrixCoefficients != AVIF_MATRIX_COEFFICIENTS_BT601 ||
        !image->alphaPlane) return 3;

    std::vector<unsigned char> output(image->width * image->height * 4);
    for (int premultiply = 0; premultiply < 2; ++premultiply) {
        // Swapping U/V with the matching constant produces byte-order RGBA.
        const int result = libyuv::I420AlphaToARGBMatrixFilter(
            image->yuvPlanes[0], image->yuvRowBytes[0],
            image->yuvPlanes[2], image->yuvRowBytes[2],
            image->yuvPlanes[1], image->yuvRowBytes[1],
            image->alphaPlane, image->alphaRowBytes,
            output.data(), image->width * 4, &libyuv::kYvuJPEGConstants,
            image->width, image->height, premultiply, libyuv::kFilterBilinear);
        if (result) return 4;
        char path[4096];
        const int count = std::snprintf(path, sizeof(path), "%s-%s.rgba", argv[2],
            premultiply ? "premultiplied" : "straight");
        if (count <= 0 || count >= static_cast<int>(sizeof(path))) return 5;
        FILE* file = std::fopen(path, "wb");
        if (!file) return 6;
        const bool written = std::fwrite(output.data(), 1, output.size(), file) == output.size();
        const bool closed = std::fclose(file) == 0;
        if (!written || !closed) return 7;
    }
    avifImageDestroy(image);
    avifDecoderDestroy(decoder);
    return 0;
}
