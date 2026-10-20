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

            try {
                page.render(
                        image,
                        null,
                        null,
                        PdfRenderer.Page.RENDER_MODE_FOR_DISPLAY
                );

                return image;
            } catch (Exception exception) {
                image.recycle();
                throw exception;
            }
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
            if (!image.compress(
                    Bitmap.CompressFormat.PNG,
                    100,
                    stream
            )) {
                throw new Exception(
                        "Could not save the page preview."
                );
            }
        } finally {
            image.recycle();
        }
    }

    public static String recognize(
            String file,
            int index,
            String dataDir
    ) throws Exception {

        File tessdata = new File(dataDir, "tessdata");

        for (String language : new String[]{"fas", "eng"}) {
            File model = new File(
                    tessdata,
                    language + ".traineddata"
            );

            if (!model.isFile() || model.length() < 100000) {
                throw new Exception(
                        "Missing or incomplete OCR language file: "
                                + language
                                + ".traineddata"
                );
            }
        }

        TessBaseAPI api = new TessBaseAPI();
        Bitmap image = null;

        try {
            // tess-two 9.1.0 requires legacy language models.
            if (!api.init(
                    dataDir,
                    "fas+eng",
                    TessBaseAPI.OEM_TESSERACT_ONLY
            )) {
                throw new Exception(
                        "OCR initialization failed. "
                                + "Use compatible legacy traineddata files."
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
            try {
                api.end();
            } finally {
                if (image != null) {
                    image.recycle();
                }
            }
        }
    }

    private PdfBridge() {
    }
}
