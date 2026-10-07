#include <cairo.h>
#include <cairo-pdf.h>
#include <stdio.h>
#include <string.h>
static void paint(cairo_t *cr, int alpha, int nested) {
    cairo_set_source_rgb(cr,1,1,1); cairo_paint(cr);
    if(nested)cairo_push_group(cr);
    cairo_push_group(cr);
    cairo_set_source_rgba(cr,1,0,0,alpha ? .5 : 1);
    cairo_rectangle(cr,10,10,60,60);cairo_fill(cr);
    cairo_rectangle(cr,30,10,60,60);cairo_fill(cr);
    cairo_pop_group_to_source(cr);cairo_paint_with_alpha(cr,.5);
    if(nested){cairo_pop_group_to_source(cr);cairo_paint_with_alpha(cr,.5);}
}
int main(int argc,char **argv) {
    if(argc!=2)return 2;
    for(int alpha=0;alpha<2;alpha++)for(int nested=0;nested<2;nested++) {
        char path[4096];snprintf(path,sizeof(path),"%s/cairo-%s%s.pdf",argv[1],alpha?"child-alpha":"opaque",nested?"-nested":"");
        cairo_surface_t *pdf=cairo_pdf_surface_create(path,100,80);
        cairo_pdf_surface_set_metadata(pdf,CAIRO_PDF_METADATA_CREATE_DATE,"2026-10-05T00:00:00Z");
        cairo_t *cr=cairo_create(pdf);paint(cr,alpha,nested);cairo_destroy(cr);cairo_surface_finish(pdf);
        if(cairo_surface_status(pdf)!=CAIRO_STATUS_SUCCESS)return 3;cairo_surface_destroy(pdf);
        cairo_surface_t *png=cairo_image_surface_create(CAIRO_FORMAT_ARGB32,100,80);
        cr=cairo_create(png);paint(cr,alpha,nested);cairo_destroy(cr);
        snprintf(path,sizeof(path),"%s/cairo-%s%s-expected.png",argv[1],alpha?"child-alpha":"opaque",nested?"-nested":"");
        if(cairo_surface_write_to_png(png,path)!=CAIRO_STATUS_SUCCESS)return 4;cairo_surface_destroy(png);
    }
    puts(cairo_version_string());return 0;
}
