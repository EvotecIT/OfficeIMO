// Development-only independent oracle bridge; Apache JBIG2 ImageIO is supplied explicitly.
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import javax.imageio.ImageIO;
import javax.imageio.stream.ImageInputStream;
import org.apache.pdfbox.jbig2.JBIG2ImageReader;
import org.apache.pdfbox.jbig2.JBIG2ImageReaderSpi;

public class DecodeJbig2 {
    public static void main(String[] args) throws Exception {
        JBIG2ImageReader reader = new JBIG2ImageReader(new JBIG2ImageReaderSpi());
        try (ImageInputStream input = ImageIO.createImageInputStream(new File(args[0]));
             OutputStream output = new FileOutputStream(args[1])) {
            reader.setInput(input);
            BufferedImage image = reader.read(0);
            output.write(("P4\n" + image.getWidth() + " " + image.getHeight() + "\n")
                .getBytes(StandardCharsets.US_ASCII));
            for (int y = 0; y < image.getHeight(); y++) {
                for (int x = 0; x < image.getWidth(); x += 8) {
                    int value = 0;
                    for (int bit = 0; bit < 8 && x + bit < image.getWidth(); bit++) {
                        if ((image.getRGB(x + bit, y) & 255) < 128) value |= 128 >> bit;
                    }
                    output.write(value);
                }
            }
        } finally {
            reader.dispose();
        }
    }
}
