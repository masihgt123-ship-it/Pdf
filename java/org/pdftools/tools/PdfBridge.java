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
            ParcelFileDescriptor descriptor =
                ParcelFileDescriptor.open(
                    new File(file),
                    ParcelFileDescriptor.MODE_READ_ONLY
                );

            PdfRenderer renderer =
                new PdfRenderer(descriptor);

            PdfRenderer.Page page =
                renderer.openPage(index)
        ) {
            float scale = Math.min(
                (float) width / page.getWidth(),
                3.0f
            );

            int renderedWidth = Math.max(
                1,
                Math.round(page.getWidth() * scale)
            );

            int renderedHeight = Math.max(
                1,
                Math.round(page.getHeight() * scale)
            );

            if (
                (long) renderedWidth * renderedHeight
                > 8000000
            ) {
                throw new Exception(
                    "Page is too large to render safely."
                );
            }

            Bitmap bitmap = Bitmap.createBitmap(
                renderedWidth,
                renderedHeight,
                Bitmap.Config.ARGB_8888
            );

            bitmap.eraseColor(Color.WHITE);

            page.render(
                bitmap,
                null,
                null,
                PdfRenderer.Page.RENDER_MODE_FOR_DISPLAY
            );

            return bitmap;
        }
    }

    public static void render(
            String file,
            int index,
            String output
    ) throws Exception {

        Bitmap bitmap = bitmap(
            file,
            index,
            1200
        );

        try (
            FileOutputStream stream =
                new FileOutputStream(output)
        ) {
            bitmap.compress(
                Bitmap.CompressFormat.PNG,
                100,
                stream
            );

        } finally {
            bitmap.recycle();
        }
    }

    public static String recognize(
            String file,
            int index,
            String dataDirectory
    ) throws Exception {

        TessBaseAPI api = new TessBaseAPI();
        Bitmap bitmap = null;

        try {
            // Tesseract engine mode 1: LSTM only.
            boolean initialized = api.init(
                dataDirectory,
                "fas+eng",
                1
            );

            if (!initialized) {
                throw new Exception(
                    "OCR initialization failed. "
                    + "Check traineddata assets."
                );
            }

            bitmap = bitmap(
                file,
                index,
                2000
            );

            api.setPageSegMode(
                TessBaseAPI.PageSegMode.PSM_AUTO
            );

            api.setImage(bitmap);

            String text = api.getUTF8Text();

            return text == null ? "" : text;

        } finally {
            api.end();

            if (bitmap != null) {
                bitmap.recycle();
            }
        }
    }

    private PdfBridge() {
    }
}
