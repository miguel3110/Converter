import converter
import flet
from flet import (
    ElevatedButton,
    FilePicker,
    FilePickerResultEvent,
    Page,
    Row,
    Text,
    icons,
)


def main(page: Page):
    page.title = ".dat CONVERTER"
    page.vertical_alignment = flet.MainAxisAlignment.CENTER

    # Save file dialog
    def save_file_result(e: FilePickerResultEvent):
        if e.path:
            save_file_path.value = converter.convert_to_excel(e.path)
            if save_file_path.value != "Invalid file path or file extension. Please select a valid .dat file.":
                open_excel_dialog.disabled = False
            else:
                open_excel_dialog.disabled = True
        else:
            save_file_path.value = ""

        save_file_path.update()
        open_excel_path.update()
        open_excel_dialog.update()

    save_file_dialog = FilePicker(on_result=save_file_result)
    save_file_path = Text("")

    def open_excel_file(e):
        open_excel_path.value = converter.open_excel_file(save_file_path.value)
        save_file_path.value = ""
        save_file_path.update()
        open_excel_path.update()

    open_excel_dialog = ElevatedButton(
        "Open file", icon=flet.icons.UPLOAD_FILE, on_click=open_excel_file, disabled=True)
    open_excel_path = Text("")
    # hide all dialogs in overlay
    page.overlay.extend(
        [save_file_dialog])

    page.add(
        Row(
            [
                ElevatedButton(
                    "Select file",
                    icon=icons.FILE_COPY,
                    on_click=lambda _: save_file_dialog.save_file(),
                    disabled=page.web,
                ),

            ],
            alignment=flet.MainAxisAlignment.CENTER
        ),

        Row(
            [
                open_excel_dialog
            ],
            alignment=flet.MainAxisAlignment.CENTER
        ),
        Row(
            [
                save_file_path,
                open_excel_path
            ],
            alignment=flet.MainAxisAlignment.CENTER
        )
    )


flet.app(target=main)
