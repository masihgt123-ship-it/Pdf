import os
import threading
from io import BytesIO

from kivy.app import App
from kivy.clock import Clock
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.utils import platform

from pypdf import PdfReader
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


MAX_BYTES = 20 * 1024 * 1024
MAX_PAGES = 300

OPEN_PDF = 101
SAVE_WORD = 102


def clean_text(text):
    return "".join(
        char for char in text
        if char in "\n\r\t"
        or 0x20 <= ord(char) <= 0xD7FF
        or 0xE000 <= ord(char) <= 0xFFFD
        or 0x10000 <= ord(char) <= 0x10FFFF
    )


def is_rtl(text):
    return any(
        "\u0590" <= char <= "\u08ff"
        or "\ufb1d" <= char <= "\ufefc"
        for char in text
    )


def convert_pdf(data, report):
    reader = PdfReader(BytesIO(data))

    if reader.is_encrypted:
        try:
            unlocked = reader.decrypt("")
        except Exception:
            unlocked = False

        if not unlocked:
            raise ValueError(
                "Password-protected PDFs are not supported."
            )

    count = len(reader.pages)

    if not count:
        raise ValueError("The PDF has no pages.")

    if count > MAX_PAGES:
        raise ValueError("Maximum supported pages: 300.")

    document = Document()
    extracted = 0
    empty = 0

    for index, page in enumerate(reader.pages):
        if index:
            document.add_page_break()

        text = clean_text(page.extract_text() or "").strip()

        if not text:
            empty += 1
            document.add_paragraph(
                "[Page %d: no extractable text]" % (index + 1)
            )
        else:
            extracted += len(text)

            for line in text.splitlines():
                paragraph = document.add_paragraph()
                rtl = is_rtl(line)

                if rtl:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    properties = paragraph._p.get_or_add_pPr()
                    bidi = OxmlElement("w:bidi")
                    bidi.set(qn("w:val"), "1")
                    properties.append(bidi)

                run = paragraph.add_run(line)
                run.font.name = "Arial"

                if rtl:
                    properties = run._r.get_or_add_rPr()
                    direction = OxmlElement("w:rtl")
                    direction.set(qn("w:val"), "1")
                    properties.append(direction)

        report("Converting page %d / %d" % (index + 1, count))

    if not extracted:
        raise ValueError(
            "No text found. Scanned PDFs need OCR."
        )

    output = BytesIO()
    document.save(output)
    return output.getvalue(), empty


class PDFWordApp(App):
    def build(self):
        self.title = "PDF to Word"
        self.result = None
        self.busy = False
        self.bound = False

        root = BoxLayout(
            orientation="vertical",
            padding=24,
            spacing=18,
        )

        root.add_widget(Label(
            text="PDF to Word",
            font_size="30sp",
            size_hint_y=0.2,
        ))

        root.add_widget(Label(
            text=(
                "Offline text conversion\n"
                "Maximum: 20 MB / 300 pages\n\n"
                "Scanned PDFs require OCR.\n"
                "Images and table layout are not preserved."
            ),
            halign="center",
            font_size="16sp",
        ))

        self.status = Label(
            text="Select a PDF to start.",
            halign="center",
            font_size="16sp",
        )
        self.status.bind(
            size=lambda widget, size: setattr(
                widget, "text_size", (size[0], None)
            )
        )
        root.add_widget(self.status)

        self.open_button = Button(
            text="Select PDF and convert",
            size_hint_y=None,
            height=64,
        )
        self.open_button.bind(on_release=self.pick_pdf)
        root.add_widget(self.open_button)

        self.save_button = Button(
            text="Save Word file",
            size_hint_y=None,
            height=64,
            disabled=True,
        )
        self.save_button.bind(on_release=self.save_word)
        root.add_widget(self.save_button)

        return root

    def on_start(self):
        if platform == "android":
            from android import activity

            activity.bind(on_activity_result=self.on_activity_result)
            self.bound = True
        else:
            self.status.text = "Run this version on Android."
            self.open_button.disabled = True

    def report(self, message):
        Clock.schedule_once(
            lambda dt: setattr(self.status, "text", message)
        )

    def set_busy(self, busy):
        self.busy = busy
        self.open_button.disabled = busy
        self.save_button.disabled = busy or self.result is None

    def launch_picker(self, save=False):
        from android.runnable import run_on_ui_thread
        from jnius import autoclass

        @run_on_ui_thread
        def launch():
            try:
                Intent = autoclass("android.content.Intent")
                PythonActivity = autoclass(
                    "org.kivy.android.PythonActivity"
                )

                action = (
                    Intent.ACTION_CREATE_DOCUMENT
                    if save else Intent.ACTION_OPEN_DOCUMENT
                )

                intent = Intent(action)
                intent.addCategory(Intent.CATEGORY_OPENABLE)

                if save:
                    intent.setType(
                        "application/vnd.openxmlformats-officedocument."
                        "wordprocessingml.document"
                    )
                    intent.putExtra(Intent.EXTRA_TITLE, "converted.docx")
                else:
                    intent.setType("application/pdf")

                PythonActivity.mActivity.startActivityForResult(
                    intent, SAVE_WORD if save else OPEN_PDF
                )
            except Exception:
                self.report("Could not open the Android file picker.")

        launch()

    def pick_pdf(self, *_):
        if platform == "android" and not self.busy:
            self.launch_picker()

    def save_word(self, *_):
        if self.result is not None and not self.busy:
            self.launch_picker(save=True)

    def on_activity_result(self, request, result_code, intent):
        if request not in (OPEN_PDF, SAVE_WORD):
            return

        # Activity.RESULT_OK == -1
        if result_code != -1 or intent is None:
            return

        uri = intent.getData()
        if uri is None:
            return

        # Keep the URI alive and dispatch UI updates to Kivy's thread.
        Clock.schedule_once(
            lambda dt: self.begin_job(request, uri)
        )

    def begin_job(self, request, uri):
        if self.busy:
            return

        if request == OPEN_PDF:
            self.result = None

        self.set_busy(True)
        self.status.text = (
            "Reading PDF..." if request == OPEN_PDF
            else "Saving Word file..."
        )

        threading.Thread(
            target=self.process_uri,
            args=(request, uri),
            daemon=True,
        ).start()

    def open_uri(self, uri, mode):
        from jnius import autoclass

        PythonActivity = autoclass(
            "org.kivy.android.PythonActivity"
        )
        resolver = (
            PythonActivity.mActivity.getContentResolver()
        )

        descriptor = resolver.openFileDescriptor(
            uri, "r" if mode == "rb" else "wt"
        )

        if descriptor is None:
            raise ValueError("Cannot access this file.")

        try:
            fd = descriptor.detachFd()
        finally:
            descriptor.close()

        try:
            return os.fdopen(fd, mode)
        except Exception:
            os.close(fd)
            raise

    def process_uri(self, request, uri):
        try:
            if request == OPEN_PDF:
                with self.open_uri(uri, "rb") as source:
                    data = source.read(MAX_BYTES + 1)

                if len(data) > MAX_BYTES:
                    raise ValueError("Maximum file size: 20 MB.")

                content, empty_pages = convert_pdf(
                    data, self.report
                )

                message = "Done! Tap Save Word file."
                if empty_pages:
                    message += (
                        "\n%d pages had no extractable text."
                        % empty_pages
                    )

                Clock.schedule_once(
                    lambda dt, c=content, m=message:
                    self.finish(c, m)
                )
            else:
                with self.open_uri(uri, "wb") as destination:
                    destination.write(self.result)

                Clock.schedule_once(
                    lambda dt: self.finish(
                        self.result, "Word file saved successfully."
                    )
                )

        except Exception as error:
            message = str(error) or "Conversion failed."
            Clock.schedule_once(
                lambda dt, m=message: self.finish(
                    self.result, "Error: " + m
                )
            )

    def finish(self, content, message):
        self.result = content
        self.status.text = message
        self.set_busy(False)

    def on_pause(self):
        return True

    def on_stop(self):
        if platform == "android" and self.bound:
            from android import activity

            activity.unbind(
                on_activity_result=self.on_activity_result
            )


if __name__ == "__main__":
    PDFWordApp().run()
