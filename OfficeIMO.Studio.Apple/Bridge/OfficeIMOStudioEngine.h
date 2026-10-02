#ifndef OFFICEIMO_STUDIO_ENGINE_H
#define OFFICEIMO_STUDIO_ENGINE_H
#include <stdint.h>
// Input is borrowed. Free every returned output, including UTF-8 error buffers.
// Operations: 0 inspect (JSON), 1 add note (PDF), 2 sample (PDF).
// Status: 0 success, -1 error text, -2 no output available. Maximum PDF: 64 MiB.
int32_t oi_studio_pdf(int32_t operation, const uint8_t *input, int32_t input_length,
                     int32_t page, const uint8_t *note, int32_t note_length,
                     uint8_t **output, int32_t *output_length);
void oi_studio_free(void *buffer);
#endif
