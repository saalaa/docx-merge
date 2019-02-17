#!/usr/bin/env python

__version__ = "1.0.1"

import csv
import os
import shutil
import tempfile
import zipfile

from jinja2 import Template, StrictUndefined
from slugify import slugify

from tkinter import (
    Tk,
    Frame,
    Text,
    Button,
    Label,
    filedialog,
    Checkbutton,
    StringVar,
    BooleanVar,
)


def cleanup(text):
    return slugify(text).upper().replace("-", "_")


def build_docx(source, destination):
    zip = zipfile.ZipFile(destination, "w")

    for root, directories, filenames in os.walk(source):
        for filename in filenames:
            abs_path = os.path.join(root, filename)
            rel_path = os.path.relpath(abs_path, source)

            zip.write(abs_path, rel_path)

    zip.close()


def docx_to_pdf(filename):
    import comtypes.client

    word = comtypes.client.CreateObject("Word.Application")

    doc = word.Documents.Open(filename)
    doc.SaveAs(filename[:-5] + ".pdf", FileFormat=17)
    doc.Close()

    word.Quit()


def merge(template, data, pattern, convert=False):
    tpl_dirname = os.path.dirname(template)
    tpl_filename = os.path.basename(template)[:-5]

    dst = tempfile.mkdtemp(prefix=tpl_filename, dir=tpl_dirname)
    tmp = os.path.join(dst, "tmp")

    document = os.path.join(tmp, "word/document.xml")

    zip = zipfile.ZipFile(template, "r")
    zip.extractall(tmp)
    zip.close()

    with open(document) as file:
        template = file.read()

    template = Template(template, undefined=StrictUndefined)
    pattern = Template(pattern, undefined=StrictUndefined)

    with open(data) as file:
        reader = csv.DictReader(file)
        for i, row in enumerate(reader):
            environment = {cleanup(k): v for k, v in row.items()}

            contents = template.render(environment)
            filename = pattern.render(environment)

            filepath = os.path.join(dst, filename)

            with open(document, "w") as file:
                file.write(contents)

            build_docx(tmp, filepath)

            if convert:
                docx_to_pdf(filepath)

    shutil.rmtree(tmp)

    return dst


class Application:
    def __init__(self):
        self.root = Tk()
        self.root.title("docx-merge %s" % __version__)

        # Template document (.docx)

        frame = Frame(self.root)
        frame.pack(side="top", fill="x", padx=5, pady=5)

        label = Label(frame, text="Template document (.docx):", justify="left")
        label.pack(side="top", fill="x")

        self.template = Text(frame, height=1)
        self.template.pack(side="left", fill="x")

        Button(frame, text="Browse", command=self.on_browse_template).pack(
            side="right"
        )

        # CSV data (.csv)

        frame = Frame(self.root)
        frame.pack(side="top", fill="x", padx=5, pady=5)

        label = Label(frame, text="CSV data (.csv):", justify="left")
        label.pack(side="top", fill="x")

        self.data = Text(frame, height=1)
        self.data.pack(side="left", fill="x")

        Button(frame, text="Browse", command=self.on_browse_data).pack(
            side="right"
        )

        # Filename pattern

        frame = Frame(self.root)
        frame.pack(side="top", fill="x", padx=5, pady=5)

        label = Label(frame, text="Filename pattern:", justify="left")
        label.pack(side="top")

        self.pattern = Text(frame, height=1)
        self.pattern.pack(side="top", fill="x")

        # --

        self.notification = StringVar()
        self.notification.set("")

        self.convert = BooleanVar()
        self.convert.set(False)

        frame = Frame(self.root)
        frame.pack(side="top", fill="x", padx=5, pady=5)

        Label(frame, textvariable=self.notification, justify="left").pack(
            side="left", fill="x"
        )

        Button(frame, text="Process", command=self.on_process).pack(
            side="right", padx=5
        )

        Checkbutton(
            frame,
            text="Convert to PDF",
            variable=self.convert,
            onvalue=True,
            offvalue=False,
        ).pack(side="right")

        # Autofill

        self.pattern.insert("insert", "{{ARCHITECTE}}-{{FACTURE}}.docx")

    def run(self):
        self.root.mainloop()

    def notify(self, text=""):
        self.notification.set(text)

    def on_browse_template(self):
        filename = filedialog.LoadFileDialog(self.root, title="Browse").go(
            pattern="*.docx", dir_or_file=os.path.expanduser("~")
        )

        if not filename:
            return

        self.template.delete("0.1", "end")
        self.template.insert("insert", filename)

        self.notify()

    def on_browse_data(self):
        filename = filedialog.LoadFileDialog(self.root, title="Browse").go(
            pattern="*.csv", dir_or_file=os.path.expanduser("~")
        )

        if not filename:
            return

        self.data.delete("0.1", "end")
        self.data.insert("insert", filename)

        self.notify()

    def on_process(self):
        template = self.template.get("1.0", "end").strip()
        data = self.data.get("1.0", "end").strip()
        pattern = self.pattern.get("1.0", "end").strip()

        convert = self.convert.get()

        if not template or not data or not pattern:
            return

        try:
            path = merge(template, data, pattern, convert=convert)
        except Exception as e:
            self.notify("Error: %s" % str(e))

            raise e

        self.notify("Done in %s" % path)


if __name__ == "__main__":
    app = Application()
    app.run()
