import copy
import re
from pathlib import Path

import wx
from docx import Document
from openpyxl import load_workbook


def safe_filename(value):
    """Make a value safe to use as a filename."""
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(value)).strip()
    return name or "document"


def replace_in_paragraph(paragraph, replacements):
    """Replace {{placeholders}} while retaining the placeholder's formatting."""
    runs = list(paragraph.runs)
    if not runs:
        return

    original_text = "".join(run.text for run in runs)
    if not original_text:
        return

    run_ranges = []
    position = 0

    for run in runs:
        start = position
        position += len(run.text)
        run_ranges.append((start, position, run))

    def run_at(char_position):
        for start, end, run in run_ranges:
            if start <= char_position < end:
                return run
        return runs[-1]

    placeholder_pattern = re.compile(r"\{\{([^{}]+)\}\}")
    matches = []

    for match in placeholder_pattern.finditer(original_text):
        key = match.group(1).strip()
        if key in replacements:
            value = replacements[key]
            replacement = "" if value is None else str(value)
            matches.append((match.start(), match.end(), replacement))

    if not matches:
        return

    segments = []

    def add_original_segments(start, end):
        for run_start, run_end, run in run_ranges:
            overlap_start = max(start, run_start)
            overlap_end = min(end, run_end)

            if overlap_start < overlap_end:
                text = original_text[overlap_start:overlap_end]
                segments.append((text, run))

    cursor = 0

    for start, end, replacement in matches:
        add_original_segments(cursor, start)

        # The replacement takes formatting from the run where the
        # placeholder starts.
        placeholder_run = run_at(start)
        if replacement:
            segments.append((replacement, placeholder_run))

        cursor = end

    add_original_segments(cursor, len(original_text))

    for run in runs:
        run._r.getparent().remove(run._r)

    for text, source_run in segments:
        new_run = paragraph.add_run(text)
        source_rpr = source_run._r.rPr

        if source_rpr is not None:
            new_run._r.insert(0, copy.deepcopy(source_rpr))


def replace_in_table(table, replacements):
    for row in table.rows:
        for cell in row.cells:
            replace_in_container(cell, replacements)


def replace_in_container(container, replacements):
    for paragraph in container.paragraphs:
        replace_in_paragraph(paragraph, replacements)

    for table in container.tables:
        replace_in_table(table, replacements)


def fill_document(template_path, replacements, output_path):
    doc = Document(template_path)

    replace_in_container(doc, replacements)

    for section in doc.sections:
        replace_in_container(section.header, replacements)
        replace_in_container(section.footer, replacements)

    doc.save(output_path)


def export_proposals(template_path, workbook_path, output_folder, sheet_name):
    workbook = load_workbook(workbook_path, data_only=True)

    if sheet_name not in workbook.sheetnames:
        raise ValueError(
            f'The selected workbook does not contain a "{sheet_name}" worksheet.'
        )

    worksheet = workbook[sheet_name]
    rows = worksheet.iter_rows(values_only=True)
    headers = next(rows, None)

    if not headers or not any(header is not None for header in headers):
        raise ValueError(f'The "{sheet_name}" worksheet has no header row.')

    header_names = [
        str(header).strip() if header is not None else None
        for header in headers
    ]

    output_dir = Path(output_folder)
    output_dir.mkdir(parents=True, exist_ok=True)
    created_count = 0

    for row_number, row in enumerate(rows, start=2):
        if not any(value is not None for value in row):
            continue

        replacements = {
            header_names[index]: row[index] if index < len(row) else ""
            for index in range(len(header_names))
            if header_names[index]
        }

        filename_value = replacements.get("FileName")
        base_name = (
            safe_filename(filename_value)
            if filename_value
            else f"proposal_{row_number}"
        )

        output_path = output_dir / f"{base_name}.docx"
        suffix = 2

        while output_path.exists():
            output_path = output_dir / f"{base_name}_{suffix}.docx"
            suffix += 1

        fill_document(template_path, replacements, output_path)
        created_count += 1

    if created_count == 0:
        raise ValueError(
            f'No proposal data rows were found beneath the headers '
            f'in "{sheet_name}".'
        )

    return created_count


class ProposalApp(wx.Frame):
    def __init__(self):
        super().__init__(
            None,
            title="Proposal Export",
            size=(720, 340),
        )

        panel = wx.Panel(self)
        layout = wx.BoxSizer(wx.VERTICAL)

        self.template_picker = wx.FilePickerCtrl(
            panel,
            message="Select the Word proposal template",
            wildcard="Word documents (*.docx)|*.docx",
            style=wx.FLP_OPEN | wx.FLP_FILE_MUST_EXIST,
        )

        self.workbook_picker = wx.FilePickerCtrl(
            panel,
            message="Select the Excel workbook",
            wildcard="Excel workbooks (*.xlsx)|*.xlsx",
            style=wx.FLP_OPEN | wx.FLP_FILE_MUST_EXIST,
        )

        self.sheet_name_ctrl = wx.TextCtrl(panel, value="proposalOutput")

        self.output_picker = wx.DirPickerCtrl(
            panel,
            message="Select the output folder",
            style=wx.DIRP_USE_TEXTCTRL,
        )

        layout.Add(
            wx.StaticText(panel, label="Word template:"),
            0,
            wx.ALL,
            5,
        )
        layout.Add(
            self.template_picker,
            0,
            wx.EXPAND | wx.LEFT | wx.RIGHT,
            5,
        )

        layout.Add(
            wx.StaticText(panel, label="Excel workbook:"),
            0,
            wx.ALL,
            5,
        )
        layout.Add(
            self.workbook_picker,
            0,
            wx.EXPAND | wx.LEFT | wx.RIGHT,
            5,
        )

        layout.Add(
            wx.StaticText(panel, label="Excel sheet name:"),
            0,
            wx.ALL,
            5,
        )
        layout.Add(
            self.sheet_name_ctrl,
            0,
            wx.EXPAND | wx.LEFT | wx.RIGHT,
            5,
        )

        layout.Add(
            wx.StaticText(panel, label="Output folder:"),
            0,
            wx.ALL,
            5,
        )
        layout.Add(
            self.output_picker,
            0,
            wx.EXPAND | wx.LEFT | wx.RIGHT,
            5,
        )

        export_button = wx.Button(panel, label="Export File")
        export_button.Bind(wx.EVT_BUTTON, self.on_export_proposal)
        layout.Add(export_button, 0, wx.ALL | wx.ALIGN_RIGHT, 10)

        panel.SetSizer(layout)
        self.Centre()

    def on_export_proposal(self, event):
        template_path = self.template_picker.GetPath()
        workbook_path = self.workbook_picker.GetPath()
        output_folder = self.output_picker.GetPath()
        sheet_name = self.sheet_name_ctrl.GetValue().strip()

        if not template_path or not workbook_path or not output_folder or not sheet_name:
            wx.MessageBox(
                "Select a Word template, Excel workbook, output folder, "
                "and enter a sheet name.",
                "Missing selection",
                wx.OK | wx.ICON_WARNING,
            )
            return

        try:
            count = export_proposals(
                template_path,
                workbook_path,
                output_folder,
                sheet_name,
            )

            wx.MessageBox(
                f"Created {count} proposal document(s) in:\n{output_folder}",
                "Export complete",
                wx.OK | wx.ICON_INFORMATION,
            )
        except Exception as error:
            wx.MessageBox(
                f"Proposal export failed:\n{error}",
                "Error",
                wx.OK | wx.ICON_ERROR,
            )


if __name__ == "__main__":
    app = wx.App(False)
    frame = ProposalApp()
    frame.Show()
    app.MainLoop()
