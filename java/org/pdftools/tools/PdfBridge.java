package org.pdftools.tools;

import android.graphics.Bitmap;
import android.graphics.Color;
import android.graphics.pdf.PdfRenderer;
import android.os.ParcelFileDescriptor;

import com.googlecode.tesseract.android.TessBaseAPI;

import java.io.File;
import java.io.FileOutputStream;


public final class PdfBridge {

    private static Bitmap bitmap(
            String file,
            int index,
            int width
    ) throws Exception {

        try (
                ParcelFileDescriptor fd =
                        ParcelFileDescriptor.open(
                                new File(file),
                                ParcelFileDescriptor.MODE_READ_ONLY
                        );

                PdfRenderer renderer = new PdfRenderer(fd);

                PdfRenderer.Page page = renderer.openPage(index)
        ) {
            float scale = Math.min(
                    (float) width / page.getWidth(),
                    3.0f
            );

            int w = Math.max(
                    1,
                    Math.round(page.getWidth() * scale)
            );

            int h = Math.max(
                    1,
                    Math.round(page.getHeight() * scale)
            );

            if ((long) w * h > 8000000) {
                throw new Exception(
                        "Page is too large to render safely."
                );
            }

            Bitmap image = Bitmap.createBitmap(
                    w,
                    h,
                    Bitmap.Config.ARGB_8888
            );

            image.eraseColor(Color.WHITE);

            page.render(
                    image,
                    null,
                    null,
                    PdfRenderer.Page.RENDER_MODE_FOR_DISPLAY
            );

            return image;
        }
    }

    public static void render(
            String file,
            int index,
            String output
    ) throws Exception {

        Bitmap image = bitmap(
                file,
                index,
                1200
        );

        try (
                FileOutputStream stream =
                        new FileOutputStream(output)
        ) {
            image.compress(
                    Bitmap.CompressFormat.PNG,
                    100,
                    stream
            );
        } finally {
            image.recycle();
        }
    }

    public static String recognize(
            String file,
            int index,
            String dataDir
    ) throws Exception {

        TessBaseAPI api = new TessBaseAPI();
        Bitmap image = null;

        try {
            // Tesseract OEM_LSTM_ONLY = 1.
            if (!api.init(dataDir, "fas+eng", 1)) {
                throw new Exception(
                        "OCR initialization failed. "
                                + "Check traineddata assets."
                );
            }

            image = bitmap(
                    file,
                    index,
                    2000
            );

            api.setPageSegMode(
                    TessBaseAPI.PageSegMode.PSM_AUTO
            );

            api.setImage(image);

            String text = api.getUTF8Text();

            return text == null ? "" : text;

        } finally {
            api.end();

            if (image != null) {
                image.recycle();
            }
        }
    }

    private PdfBridge() {
    }
}
