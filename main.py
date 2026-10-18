import os
import re
import shutil
import threading
from io import BytesIO

from kivy.app import App
from kivy.clock import Clock
from kivy.core.image import Image as CoreImage
from kivy.graphics import Color, Line
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.image import Image
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.uix.textinput import TextInput
from kivy.uix.widget import Widget
from kivy.utils import platform

from pypdf import PdfReader, PdfWriter
from pypdf.generic import (
    DictionaryObject,
    NameObject,
    DecodedStreamObject,
)

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


MAX_BYTES = 20 * 1024 * 1024
MAX_PAGES = 300

DOCX = (
    'application/vnd.openxmlformats-officedocument.'
    'wordprocessingml.document'
)


def reader_for(data):
    reader = PdfReader(BytesIO(data))

    if reader.is_encrypted:
        raise ValueError(
            'Encrypted PDFs are not supported. '
            'Open an unencrypted copy.'
        )

    if not 0 < len(reader.pages) <= MAX_PAGES:
        raise ValueError('Supported page count: 1 to 300.')

    return reader


def pdf_bytes(writer):
    output = BytesIO()
    writer.write(output)
    return output.getvalue()


def clean_text(text):
    return ''.join(
        character
        for character in text
        if (
            character in '\n\r\t'
            or 0x20 <= ord(character) <= 0xD7FF
            or 0xE000 <= ord(character) <= 0xFFFD
            or 0x10000 <= ord(character) <= 0x10FFFF
        )
    )


def word_bytes(texts):
    document = Document()

    for index, text in enumerate(texts):
        if index:
            document.add_page_break()

        cleaned = (
            clean_text(text).strip()
            or '[No text found on this page]'
        )

        for line in cleaned.splitlines():
            paragraph = document.add_paragraph()

            rtl = any(
                '\u0590' <= character <= '\u08ff'
                or '\ufb1d' <= character <= '\ufefc'
                for character in line
            )

            if rtl:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT

                tag = OxmlElement('w:bidi')
                tag.set(qn('w:val'), '1')
                paragraph._p.get_or_add_pPr().append(tag)

            run = paragraph.add_run(line)
            run.font.name = 'Arial'

            if rtl:
                tag = OxmlElement('w:rtl')
                tag.set(qn('w:val'), '1')
                run._r.get_or_add_rPr().append(tag)

    output = BytesIO()
    document.save(output)

    return output.getvalue()


def page_numbers(value, count):
    result = []

    for part in value.replace(' ', '').split(','):
        if not re.fullmatch(r'\d+(-\d+)?', part):
            raise ValueError(
                'Use page numbers like 1-3,5 (English digits).'
            )

        ends = [int(item) for item in part.split('-')]
        first, last = ends[0], ends[-1]

        if not 1 <= first <= last <= count:
            raise ValueError(
                'Page range is outside this document.'
            )

        for number in range(first - 1, last):
            if number not in result:
                result.append(number)

    return result


class SignaturePad(Widget):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)

        self.strokes = []
        self.active = None

    def normal(self, touch):
        return (
            max(
                0,
                min(
                    1,
                    (touch.x - self.x) / max(1, self.width),
                ),
            ),
            max(
                0,
                min(
                    1,
                    (touch.y - self.y) / max(1, self.height),
                ),
            ),
        )

    def on_touch_down(self, touch):
        if (
            not self.collide_point(*touch.pos)
            or self.active is not None
        ):
            return super().on_touch_down(touch)

        touch.grab(self)
        self.active = touch.uid

        self.strokes.append([self.normal(touch)])

        with self.canvas:
            Color(0.2, 0.7, 1, 1)
            self.line = Line(
                points=list(touch.pos),
                width=2,
            )

        return True

    def on_touch_move(self, touch):
        if touch.grab_current is self:
            self.strokes[-1].append(self.normal(touch))
            self.line.points += list(touch.pos)
            return True

        return super().on_touch_move(touch)

    def on_touch_up(self, touch):
        if touch.grab_current is self:
            touch.ungrab(self)
            self.active = None
            return True

        return super().on_touch_up(touch)

    def clear(self):
        self.canvas.clear()
        self.strokes = []


class PDFToolsApp(App):
    def build(self):
        self.title = 'PDF Tools'

        self.busy = False
        self.picker_pending = False
        self.bound = False

        self.data = None
        self.page = 0
        self.result = None
        self.pending_save = None
        self.buttons = []

        root = BoxLayout(
            orientation='vertical',
            padding=10,
            spacing=6,
        )

        root.add_widget(
            Label(
                text='PDF Tools - Offline',
                size_hint_y=None,
                height=35,
            )
        )

        self.preview = Image()
        root.add_widget(self.preview)

        self.counter = Label(
            text='No PDF selected',
            size_hint_y=None,
            height=28,
        )
        root.add_widget(self.counter)

        self.add_row(
            root,
            [
                ('Open PDF', lambda: self.pick(101)),
                ('Previous', lambda: self.navigate(-1)),
                ('Next', lambda: self.navigate(1)),
            ],
        )

        self.add_row(
            root,
            [
                ('Text to Word', lambda: self.convert(False)),
                ('OCR to Word', lambda: self.convert(True)),
            ],
        )

        self.add_row(
            root,
            [
                (
                    'Merge PDF',
                    lambda: (
                        self.pick(102)
                        if self.require()
                        else None
                    ),
                ),
                ('Extract pages', self.extract),
            ],
        )

        self.add_row(
            root,
            [
                ('Add note', self.note),
                ('Draw signature', self.signature),
            ],
        )

        self.add_row(
            root,
            [
                ('Save result', self.save),
            ],
        )

        self.status = Label(
            text='Open a PDF. Limit: 20 MB / 300 pages.',
            size_hint_y=None,
            height=72,
            halign='center',
        )

        self.status.bind(
            size=lambda widget, size: setattr(
                widget,
                'text_size',
                (size[0], None),
            )
        )

        root.add_widget(self.status)

        self.local_pdf = os.path.join(
            self.user_data_dir,
            'current.pdf',
        )

        self.ocr_dir = os.path.join(
            self.user_data_dir,
            'ocr',
        )

        return root

    def add_row(self, root, actions):
        row = BoxLayout(
            spacing=6,
            size_hint_y=None,
            height=48,
        )

        for title, action in actions:
            button = Button(
                text=title,
                font_size='13sp',
            )

            button.bind(
                on_release=lambda _, fn=action: fn()
            )

            self.buttons.append(button)
            row.add_widget(button)

        root.add_widget(row)

    def lock(self):
        for button in self.buttons:
            button.disabled = (
                self.busy or self.picker_pending
            )

    def report(self, text):
        Clock.schedule_once(
            lambda dt: setattr(self.status, 'text', text)
        )

    def require(self):
        if self.data is None:
            self.status.text = 'Open a PDF first.'
            return False

        return True

    def job(self, function):
        if self.busy:
            return

        self.busy = True
        self.lock()
        self.status.text = 'Working...'

        def work():
            try:
                update = function()

                Clock.schedule_once(
                    lambda dt: done(update, None)
                )

            except Exception as error:
                Clock.schedule_once(
                    lambda dt, message=str(error): done(
                        None,
                        message,
                    )
                )

        def done(update, error):
            try:
                if error:
                    self.status.text = 'Error: ' + error

                elif update:
                    update()

            except Exception as error:
                self.status.text = 'Error: ' + str(error)

            finally:
                self.busy = False
                self.lock()

        threading.Thread(
            target=work,
            daemon=True,
        ).start()

    def on_start(self):
        if platform == 'android':
            from android import activity

            activity.bind(
                on_activity_result=self.activity_result
            )
            self.bound = True

        else:
            self.status.text = 'This build requires Android.'

            for button in self.buttons:
                button.disabled = True

    def pick(self, request):
        if self.busy or self.picker_pending:
            return

        from android.runnable import run_on_ui_thread
        from jnius import autoclass

        self.picker_pending = True
        self.lock()

        @run_on_ui_thread
        def launch():
            try:
                Intent = autoclass(
                    'android.content.Intent'
                )

                activity = autoclass(
                    'org.kivy.android.PythonActivity'
                ).mActivity

                action = (
                    Intent.ACTION_CREATE_DOCUMENT
                    if request == 103
                    else Intent.ACTION_OPEN_DOCUMENT
                )

                intent = Intent(action)
                intent.addCategory(
                    Intent.CATEGORY_OPENABLE
                )

                mime_type = (
                    self.pending_save[1]
                    if request == 103
                    else 'application/pdf'
                )

                intent.setType(mime_type)

                if request == 103:
                    intent.putExtra(
                        Intent.EXTRA_TITLE,
                        self.pending_save[2],
                    )

                activity.startActivityForResult(
                    intent,
                    request,
                )

            except Exception as error:
                Clock.schedule_once(
                    lambda dt, message=str(error):
                    self.picker_failed(message)
                )

        launch()

    def picker_failed(self, message):
        self.picker_pending = False
        self.lock()

        self.status.text = (
            'File picker error: ' + message
        )

    def activity_result(self, request, code, intent):
        if request not in (101, 102, 103):
            return

        uri = (
            intent.getData()
            if code == -1 and intent is not None
            else None
        )

        Clock.schedule_once(
            lambda dt: self.selected(request, uri)
        )

    def open_uri(self, uri, mode):
        from jnius import autoclass

        resolver = autoclass(
            'org.kivy.android.PythonActivity'
        ).mActivity.getContentResolver()

        descriptor = resolver.openFileDescriptor(
            uri,
            'r' if mode == 'rb' else 'wt',
        )

        if descriptor is None:
            raise ValueError(
                'Cannot access the selected file.'
            )

        try:
            file_descriptor = descriptor.detachFd()

        finally:
            descriptor.close()

        try:
            return os.fdopen(file_descriptor, mode)

        except Exception:
            os.close(file_descriptor)
            raise

    def selected(self, request, uri):
        self.picker_pending = False
        self.lock()

        if uri is None:
            self.status.text = 'Selection cancelled.'
            return

        def work():
            if request == 103:
                payload = self.pending_save[0]

                with self.open_uri(uri, 'wb') as stream:
                    stream.write(payload)

                return lambda: setattr(
                    self.status,
                    'text',
                    'File saved successfully.',
                )

            with self.open_uri(uri, 'rb') as stream:
                data = stream.read(MAX_BYTES + 1)

            if len(data) > MAX_BYTES:
                raise ValueError(
                    'Maximum input file size: 20 MB.'
                )

            incoming = reader_for(data)

            if request == 102:
                current = reader_for(self.data)

                total_pages = (
                    len(current.pages)
                    + len(incoming.pages)
                )

                if total_pages > MAX_PAGES:
                    raise ValueError(
                        'Merged PDF cannot exceed 300 pages.'
                    )

                writer = PdfWriter()

                for reader in (current, incoming):
                    for page in reader.pages:
                        writer.add_page(page)

                data = pdf_bytes(writer)

            message = (
                'PDF opened.'
                if request == 101
                else 'Merged. Save result to export.'
            )

            return self.prepare_document(
                data,
                message,
                export=request == 102,
            )

        self.job(work)

    def bridge(self):
        from jnius import autoclass

        return autoclass(
            'org.pdftools.tools.PdfBridge'
        )

    def prepare_document(
        self,
        data,
        message,
        export=True,
        page=0,
    ):
        count = len(reader_for(data).pages)
        page = min(page, count - 1)

        pending = self.local_pdf + '.pending'

        try:
            with open(pending, 'wb') as stream:
                stream.write(data)

            image_data = self.render(page, pending)
            os.replace(pending, self.local_pdf)

        finally:
            if os.path.exists(pending):
                os.remove(pending)

        def update():
            self.data = data
            self.page = page
            self.count = count

            self.result = (
                (data, 'application/pdf', 'edited.pdf')
                if export
                else None
            )

            self.set_preview(image_data)

            self.counter.text = (
                'Page %d / %d' % (page + 1, count)
            )

            self.status.text = message

        return update

    def render(self, page, source=None):
        path = os.path.join(
            self.user_data_dir,
            'preview.png',
        )

        self.bridge().render(
            source or self.local_pdf,
            page,
            path,
        )

        with open(path, 'rb') as stream:
            return stream.read()

    def set_preview(self, data):
        self.preview.texture = CoreImage(
            BytesIO(data),
            ext='png',
        ).texture

    def navigate(self, delta):
        if not self.require():
            return

        page = max(
            0,
            min(
                self.count - 1,
                self.page + delta,
            ),
        )

        def work():
            image_data = self.render(page)

            def update():
                self.page = page
                self.set_preview(image_data)

                self.counter.text = (
                    'Page %d / %d'
                    % (page + 1, self.count)
                )

                self.status.text = 'Ready.'

            return update

        self.job(work)

    def install_languages(self):
        destination = os.path.join(
            self.ocr_dir,
            'tessdata',
        )

        os.makedirs(destination, exist_ok=True)

        assets = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            'tessdata',
        )

        for language in ('fas', 'eng'):
            target = os.path.join(
                destination,
                language + '.traineddata',
            )

            if not os.path.exists(target):
                shutil.copyfile(
                    os.path.join(
                        assets,
                        language + '.traineddata',
                    ),
                    target,
                )

    def convert(self, ocr):
        if not self.require():
            return

        def work():
            reader = reader_for(self.data)

            if ocr:
                self.install_languages()

            texts = []

            for index, page in enumerate(reader.pages):
                self.report(
                    ('OCR' if ocr else 'Extracting')
                    + ' page %d / %d'
                    % (index + 1, len(reader.pages))
                )

                if ocr:
                    text = self.bridge().recognize(
                        self.local_pdf,
                        index,
                        self.ocr_dir,
                    )

                else:
                    text = page.extract_text() or ''

                texts.append(text)

            if not any(text.strip() for text in texts):
                raise ValueError(
                    'No text found. '
                    'Try OCR for scanned PDFs.'
                )

            content = word_bytes(texts)

            empty_pages = sum(
                not text.strip()
                for text in texts
            )

            def update():
                self.result = (
                    content,
                    DOCX,
                    'converted.docx',
                )

                self.status.text = (
                    'Word ready. Save result. Empty pages: %d'
                    % empty_pages
                )

            return update

        self.job(work)

    def ask(self, title, hint, callback):
        box = BoxLayout(
            orientation='vertical',
            spacing=10,
            padding=10,
        )

        entry = TextInput(
            hint_text=hint,
            multiline=False,
        )

        box.add_widget(entry)

        row = BoxLayout(
            size_hint_y=None,
            height=50,
        )

        box.add_widget(row)

        popup = Popup(
            title=title,
            content=box,
            size_hint=(0.95, 0.45),
        )

        def apply():
            popup.dismiss()
            callback(entry.text)

        for text, action in [
            ('Cancel', popup.dismiss),
            ('Apply', apply),
        ]:
            button = Button(text=text)

            button.bind(
                on_release=lambda _, fn=action: fn()
            )

            row.add_widget(button)

        popup.open()

    def extract(self):
        if self.require():
            self.ask(
                'Extract pages into a new PDF',
                'Example: 1-3,5',
                self.extract_range,
            )

    def extract_range(self, value):
        def work():
            reader = reader_for(self.data)
            writer = PdfWriter()

            for index in page_numbers(
                value,
                len(reader.pages),
            ):
                writer.add_page(reader.pages[index])

            return self.prepare_document(
                pdf_bytes(writer),
                'Extracted pages. Save result.',
            )

        self.job(work)

    def note(self):
        if self.require():
            self.ask(
                'Visible note at top (English only)',
                'Short English note, up to 80 characters',
                self.add_note,
            )

    def add_note(self, text):
        invalid = (
            not text.strip()
            or len(text) > 80
            or any(
                not 32 <= ord(character) <= 126
                for character in text
            )
        )

        if invalid:
            self.status.text = (
                'Use 1-80 printable English characters '
                'for PDF notes.'
            )
            return

        self.stamp(note=text)

    def signature(self):
        if not self.require():
            return

        box = BoxLayout(
            orientation='vertical',
            spacing=8,
            padding=10,
        )

        box.add_widget(
            Label(
                text=(
                    'Draw here. Placed at bottom-right '
                    'of current page.'
                ),
                size_hint_y=None,
                height=45,
            )
        )

        pad = SignaturePad()
        box.add_widget(pad)

        row = BoxLayout(
            size_hint_y=None,
            height=50,
        )

        box.add_widget(row)

        popup = Popup(
            title=(
                'Handwritten signature '
                '(not a digital certificate)'
            ),
            content=box,
            size_hint=(0.98, 0.8),
        )

        def apply():
            strokes = [
                stroke[:]
                for stroke in pad.strokes
                if len(stroke) > 1
            ]

            if not strokes:
                return

            popup.dismiss()
            self.stamp(strokes=strokes)

        for text, action in [
            ('Cancel', popup.dismiss),
            ('Clear', pad.clear),
            ('Apply', apply),
        ]:
            button = Button(text=text)

            button.bind(
                on_release=lambda _, fn=action: fn()
            )

            row.add_widget(button)

        popup.open()

    def stamp(self, note=None, strokes=None):
        def work():
            reader = reader_for(self.data)
            page = reader.pages[self.page]

            page.transfer_rotation_to_content()

            left = float(page.mediabox.left)
            bottom = float(page.mediabox.bottom)
            width = float(page.mediabox.width)
            height = float(page.mediabox.height)

            overlay_writer = PdfWriter()

            overlay = overlay_writer.add_blank_page(
                width=width,
                height=height,
            )

            commands = [
                'q',
                '0 0 0 RG',
                '0 0 0 rg',
                '1.5 w',
            ]

            if note:
                escaped = (
                    note
                    .replace('\\', '\\\\')
                    .replace('(', '\\(')
                    .replace(')', '\\)')
                )

                font = DictionaryObject(
                    {
                        NameObject('/Type'):
                            NameObject('/Font'),

                        NameObject('/Subtype'):
                            NameObject('/Type1'),

                        NameObject('/BaseFont'):
                            NameObject('/Helvetica'),
                    }
                )

                overlay[NameObject('/Resources')] = (
                    DictionaryObject(
                        {
                            NameObject('/Font'):
                                DictionaryObject(
                                    {
                                        NameObject('/F1'):
                                            font,
                                    }
                                ),
                        }
                    )
                )

                font_size = max(
                    1,
                    min(
                        12,
                        (width - 30)
                        / max(1, len(note) * 0.65),
                    ),
                )

                commands.append(
                    'BT /F1 %.2f Tf '
                    '15 %.2f Td (%s) Tj ET'
                    % (
                        font_size,
                        max(5, height - 25),
                        escaped,
                    )
                )

            if strokes:
                signature_width = min(
                    180,
                    width * 0.6,
                )

                signature_height = min(
                    70,
                    height * 0.2,
                )

                x = max(
                    0,
                    width - signature_width - 15,
                )

                y = min(
                    20,
                    height * 0.05,
                )

                for stroke in strokes:
                    for index, (px, py) in enumerate(stroke):
                        commands.append(
                            '%.2f %.2f %s'
                            % (
                                x + px * signature_width,
                                y + py * signature_height,
                                'm' if index == 0 else 'l',
                            )
                        )

                    commands.append('S')

            commands.append('Q')

            stream = DecodedStreamObject()

            stream.set_data(
                '\n'.join(commands).encode('ascii')
            )

            overlay[NameObject('/Contents')] = stream

            from pypdf import Transformation

            page.merge_transformed_page(
                overlay,
                Transformation().translate(
                    left,
                    bottom,
                ),
            )

            writer = PdfWriter()

            for document_page in reader.pages:
                writer.add_page(document_page)

            return self.prepare_document(
                pdf_bytes(writer),
                (
                    'Applied. Save result. '
                    'Original file is unchanged.'
                ),
                page=self.page,
            )

        self.job(work)

    def save(self):
        if self.result is None:
            self.status.text = (
                'Create a Word file or '
                'edit/merge/extract PDF first.'
            )
            return

        self.pending_save = self.result
        self.pick(103)

    def on_pause(self):
        return True

    def on_stop(self):
        if platform == 'android' and self.bound:
            from android import activity

            activity.unbind(
                on_activity_result=self.activity_result
            )


if __name__ == '__main__':
    PDFToolsApp().run()
