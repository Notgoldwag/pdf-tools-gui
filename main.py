import sys

sys.path.append('C:\\Users\\Jaisharan A\\Downloads\\PyCharm project files')
from PDF_Tools.main import Pdf_Tools
import PySimpleGUI as sg

tool = Pdf_Tools()


def main_menu():
    main_label = sg.Text("PDF TOOLS", font=("Helvetica", 20, 'bold'), justification='center')
    split_pdf_button = sg.Button('Split PDF')
    merge_pdf_button = sg.Button('Merge PDF')
    extract_text_button = sg.Button('Extract Text')
    text_to_speech_button = sg.Button('Text to Speech')
    protect_pdf_button = sg.Button('Protect PDF')
    compress_pdf_button = sg.Button('Compress PDF')
    convert_to_png_button = sg.Button('Convert to PNG')
    convert_to_docx_button = sg.Button('Convert to DOCX')
    summarize_text_button = sg.Button('Summarize Text')
    spacer = sg.Text("")
    exit_button = sg.Button('Exit')

    layout = [
        [main_label],
        [split_pdf_button, merge_pdf_button],
        [text_to_speech_button, extract_text_button],
        [protect_pdf_button, compress_pdf_button],
        [convert_to_png_button, convert_to_docx_button],
        [summarize_text_button],
        [spacer],
        [exit_button]
    ]

    window = sg.Window('PDF Tools', layout, size=(400, 350), font=("Times New Roman", 15))

    while True:
        event, value = window.read()

        if event == sg.WIN_CLOSED or event == 'Exit':
            break
        else:
            open_tool_window(event)

    window.close()


def open_tool_window(tool_name):
    window = None

    if tool_name == 'Merge PDF':
        merge_pdf_menu_label = sg.Text('Merge PDF Tool', font=("Helvetica", 20, 'bold'), justification='center')

        merge_pdf_input_text = sg.Text("Please select the first PDF    ")
        merge_pdf_input = sg.Input(disabled=True)
        merge_pdf_select_button = sg.FilesBrowse('Choose')

        merge_pdf_spacer = sg.Text("")

        merge_pdf_input_text2 = sg.Text("Please select the second PDF")
        merge_pdf_input2 = sg.Input(disabled=True)
        merge_pdf_select_button2 = sg.FilesBrowse("Choose")

        merge_pdf_input_text3 = sg.Text("Please choose Destination Folder")
        merge_pdf_input3 = sg.Input(disabled=True)
        merge_pdf_select_button3 = sg.FolderBrowse("Choose")

        merge_button = sg.Button("Merge")

        merge_pdf_layout = [
            [merge_pdf_menu_label],
            [merge_pdf_spacer],
            [merge_pdf_input_text, merge_pdf_input, merge_pdf_select_button],
            [merge_pdf_input_text2, merge_pdf_input2, merge_pdf_select_button2],
            [merge_pdf_input_text3, merge_pdf_input3, merge_pdf_select_button3],
            [sg.Text('')],
            [merge_button],
            [sg.Text('')],
            [sg.Text("", key='-MERGE-MSG-')]
        ]
        window = sg.Window(tool_name, merge_pdf_layout, size=(850, 350), font=("Times New Roman", 15))

        while True:
            event, values = window.read()

            if event == sg.WIN_CLOSED:
                break
            elif event == "Merge":
                inp_path = merge_pdf_layout[2][1].get()
                inp_path2 = merge_pdf_layout[3][1].get()
                out_path = merge_pdf_layout[4][1].get()
                tool.merge_pdf(inp_path, inp_path2, out_path)
                window['-MERGE-MSG-'].update("PDFs merged successfully!")




    elif tool_name == 'Split PDF':
        split_pdf_menu_label = sg.Text('Split PDF Tool', font=("Helvetica", 20, 'bold'), justification='center')

        split_pdf_input_text = sg.Text("Please select the PDF file to split")
        split_pdf_input = sg.Input(disabled=True)
        split_pdf_select_button = sg.FilesBrowse('Choose')

        split_pdf_spacer = sg.Text("")

        split_pdf_input_text2 = sg.Text("Please enter the start page number")
        split_pdf_input2 = sg.InputText(key='start_page')

        split_pdf_input_text3 = sg.Text("Please enter the end page number")
        split_pdf_input3 = sg.InputText(key='end_page')

        split_pdf_input_text4 = sg.Text("Please choose Destination Folder")
        split_pdf_input4 = sg.Input(disabled=True)
        split_pdf_select_button4 = sg.FolderBrowse("Choose")

        split_button = sg.Button("Split")

        split_pdf_layout = [
            [split_pdf_menu_label],
            [split_pdf_spacer],
            [split_pdf_input_text, split_pdf_input, split_pdf_select_button],
            [split_pdf_input_text2, split_pdf_input2],
            [split_pdf_input_text3, split_pdf_input3],
            [split_pdf_input_text4, split_pdf_input4, split_pdf_select_button4],
            [sg.Text('')],
            [split_button],
            [sg.Text("", key='-SPLIT-MSG-')]
        ]

        window = sg.Window(tool_name, split_pdf_layout, size=(850, 350), font=("Times New Roman", 15))

        while True:
            event, values = window.read()

            if event == sg.WIN_CLOSED:
                break
            elif event == "Split":
                inp_path = split_pdf_layout[2][1].get()
                start_p = split_pdf_layout[3][1].get()
                end_p = split_pdf_layout[4][1].get()
                des_path = split_pdf_layout[5][1].get()
                tool.cut_pdf(inp_path, int(start_p), int(end_p), des_path)
                window['-SPLIT-MSG-'].update("PDF split successfully!")


    elif tool_name == 'Protect PDF':
        protect_pdf_menu_label = sg.Text('Protect PDF Tool', font=("Helvetica", 20, 'bold'), justification='center')

        protect_pdf_input_text = sg.Text("Please select the PDF file to protect")
        protect_pdf_input = sg.Input(disabled=True)
        protect_pdf_select_button = sg.FilesBrowse('Choose')

        protect_pdf_spacer = sg.Text("")

        protect_pdf_input_text2 = sg.Text("Please enter the password")
        protect_pdf_input2 = sg.InputText(key='password')

        protect_pdf_button = sg.Button("Protect")

        protect_pdf_layout = [
            [protect_pdf_menu_label],
            [protect_pdf_spacer],
            [protect_pdf_input_text, protect_pdf_input, protect_pdf_select_button],
            [protect_pdf_input_text2, protect_pdf_input2],
            [sg.Text('')],
            [protect_pdf_button],
            [sg.Text("", key='-PROTECT-MSG-')]
        ]

        window = sg.Window(tool_name, protect_pdf_layout, size=(900, 300), font=("Times New Roman", 15))

        while True:
            event, values = window.read()

            if event == sg.WIN_CLOSED:
                break
            elif event == "Protect":
                inp_path = protect_pdf_layout[2][1].get()
                password = protect_pdf_layout[3][1].get()
                tool.protect_pdf(inp_path, password)
                window['-PROTECT-MSG-'].update("PDF protected successfully!")


    elif tool_name == 'Extract Text':
        extract_text_menu_label = sg.Text('Extract Text Tool', font=("Helvetica", 20, 'bold'), justification='center')

        extract_text_input_text = sg.Text("Please select the PDF file to extract text from")
        extract_text_input = sg.Input(disabled=True)
        extract_text_select_button = sg.FilesBrowse('Choose')

        extract_text_spacer = sg.Text("")

        extract_text_button = sg.Button("Extract Text")

        extract_text_output_text = sg.Text("Extracted Text:")
        extract_text_output = sg.Multiline(size=(90, 40), key='-EXTRACTED-TEXT-', disabled=True, autoscroll=True)

        extract_text_layout = [
            [extract_text_menu_label],
            [extract_text_spacer],
            [extract_text_input_text, extract_text_input, extract_text_select_button],
            [sg.Text('')],
            [extract_text_button],
            [sg.Text('')],
            [extract_text_output_text],
            [extract_text_output],
            [sg.Text("", key='-EXTRACT-MSG-')]
        ]

        window = sg.Window(tool_name, extract_text_layout, size=(950, 500), font=("Times New Roman", 15))

        while True:
            event, values = window.read()

            if event == sg.WIN_CLOSED:
                break
            elif event == "Extract Text":
                input_pdf_filepath = extract_text_layout[2][1].get()
                extracted_text = tool.extract_text(input_pdf_filepath)
                window['-EXTRACTED-TEXT-'].update(extracted_text)

        window.close()


    elif tool_name == 'Convert to DOCX':
        convert_to_docx_menu_label = sg.Text('Convert to DOCX Tool', font=("Helvetica", 20, 'bold'),
                                             justification='center')

        convert_to_docx_input_text = sg.Text("Please select the PDF file to convert")
        convert_to_docx_input = sg.Input(disabled=True)
        convert_to_docx_select_button = sg.FilesBrowse('Choose')

        convert_to_docx_input_text2 = sg.Text("Please choose the output folder")
        convert_to_docx_input2 = sg.Input(disabled=True)
        convert_to_docx_select_button2 = sg.FolderBrowse('Choose')

        convert_to_docx_button = sg.Button("Convert")

        convert_to_docx_layout = [
            [convert_to_docx_menu_label],
            [convert_to_docx_input_text, convert_to_docx_input, convert_to_docx_select_button],
            [convert_to_docx_input_text2, convert_to_docx_input2, convert_to_docx_select_button2],
            [sg.Text('')],
            [convert_to_docx_button],
            [sg.Text("", key='-CONVERT-MSG-')]
        ]

        window = sg.Window(tool_name, convert_to_docx_layout, size=(900, 280), font=("Times New Roman", 15))

        while True:
            event, values = window.read()

            if event == sg.WIN_CLOSED:
                break
            elif event == "Convert":
                input_pdf_filepath = convert_to_docx_layout[1][1].get()
                output_folder = convert_to_docx_layout[2][1].get()
                tool.pdf_to_docx(input_pdf_filepath, output_folder)
                window['-CONVERT-MSG-'].update("PDF converted to DOCX successfully!")

    elif tool_name == 'Convert to PNG':
        convert_to_png_menu_label = sg.Text('Convert to PNG Tool', font=("Helvetica", 20, 'bold'),
                                            justification='center')

        convert_to_png_input_text = sg.Text("Please select the PDF file to convert")
        convert_to_png_input = sg.Input(disabled=True)
        convert_to_png_select_button = sg.FilesBrowse('Choose')

        convert_to_png_input_text2 = sg.Text("Please select the output folder")
        convert_to_png_input2 = sg.Input(disabled=True)
        convert_to_png_select_button2 = sg.FolderBrowse("Choose")

        convert_to_png_button = sg.Button("Convert")

        convert_to_png_layout = [
            [convert_to_png_menu_label],
            [sg.Text('')],
            [convert_to_png_input_text, convert_to_png_input, convert_to_png_select_button],
            [convert_to_png_input_text2, convert_to_png_input2, convert_to_png_select_button2],
            [sg.Text('')],
            [convert_to_png_button],
            [sg.Text("", key='-CONVERT-TO-PNG-MSG-')]
        ]

        window = sg.Window(tool_name, convert_to_png_layout, size=(900, 300), font=("Times New Roman", 15))

        while True:
            event, values = window.read()

            if event == sg.WIN_CLOSED:
                break
            elif event == "Convert":
                input_pdf_filepath = convert_to_png_layout[2][1].get()
                output_folder = convert_to_png_layout[3][1].get()
                tool.pdf_to_png(input_pdf_filepath, output_folder)
                window['-CONVERT-TO-PNG-MSG-'].update("PDF converted to PNG successfully!")

    elif tool_name == 'Compress PDF':
        compress_pdf_menu_label = sg.Text('Compress PDF Tool', font=("Helvetica", 20, 'bold'), justification='center')

        compress_pdf_input_text = sg.Text("Please select the PDF file to compress")
        compress_pdf_input = sg.Input(disabled=True)
        compress_pdf_select_button = sg.FilesBrowse('Choose')

        compress_pdf_spacer = sg.Text("")

        compress_pdf_input_text2 = sg.Text("Please choose Destination Folder")
        compress_pdf_input2 = sg.Input(disabled=True)
        compress_pdf_select_button2 = sg.FolderBrowse("Choose")

        compress_button = sg.Button("Compress")

        compress_pdf_layout = [
            [compress_pdf_menu_label],
            [compress_pdf_spacer],
            [compress_pdf_input_text, compress_pdf_input, compress_pdf_select_button],
            [compress_pdf_input_text2, compress_pdf_input2, compress_pdf_select_button2],
            [sg.Text('')],
            [compress_button],
            [sg.Text("", key='-COMPRESS-MSG-')]
        ]

        window = sg.Window(tool_name, compress_pdf_layout, size=(900, 300), font=("Times New Roman", 15))

        while True:
            event, values = window.read()

            if event == sg.WIN_CLOSED:
                break
            elif event == "Compress":
                input_pdf_filepath = compress_pdf_layout[2][1].get()
                output_folder = compress_pdf_layout[3][1].get()
                tool.compress(input_pdf_filepath, output_folder)
                window['-COMPRESS-MSG-'].update("PDF compressed and saved successfully!")

    elif tool_name == 'Text to Speech':
        text_to_speech_menu_label = sg.Text('Text to Speech Tool', font=("Helvetica", 20, 'bold'),
                                            justification='center')

        text_to_speech_input_text = sg.Text("Please select the text file to convert to speech")
        text_to_speech_input = sg.Input(disabled=True)
        text_to_speech_select_button = sg.FilesBrowse('Choose')

        text_to_speech_spacer = sg.Text("")

        text_to_speech_input_text2 = sg.Text("Please choose Destination Folder")
        text_to_speech_input2 = sg.Input(disabled=True)
        text_to_speech_select_button2 = sg.FolderBrowse("Choose")

        text_to_speech_button = sg.Button("Convert to Speech")

        text_to_speech_layout = [
            [text_to_speech_menu_label],
            [text_to_speech_spacer],
            [text_to_speech_input_text, text_to_speech_input, text_to_speech_select_button],
            [text_to_speech_input_text2, text_to_speech_input2, text_to_speech_select_button2],
            [sg.Text('')],
            [text_to_speech_button],
            [sg.Text("", key='-TEXTTOSPEECH-MSG-')]
        ]

        window = sg.Window(tool_name, text_to_speech_layout, size=(950, 325), font=("Times New Roman", 15))

        while True:
            event, values = window.read()

            if event == sg.WIN_CLOSED:
                break
            elif event == "Convert to Speech":
                input_text_filepath = text_to_speech_layout[2][1].get()
                output_folder = text_to_speech_layout[3][1].get()
                give = tool.extract_text(input_text_filepath)
                tool.text_to_speech(give, output_folder)
                window['-TEXTTOSPEECH-MSG-'].update("Text converted to speech and saved successfully!")

    elif tool_name == 'Summarize Text':
        summarize_text_menu_label = sg.Text('Summarize Text Tool', font=("Helvetica", 20, 'bold'),
                                            justification='center')

        summarize_text_input_text = sg.Text("Please select the text file to summarize")
        summarize_text_input = sg.Input(disabled=True)
        summarize_text_select_button = sg.FilesBrowse('Choose')

        summarize_text_spacer = sg.Text("")

        summarize_text_button = sg.Button("Summarize Text")

        summarize_text_output_text = sg.Text("Summarized Text:")
        summarize_text_output = sg.Multiline(size=(90, 40), key='-SUMMARIZED-TEXT-', disabled=True, autoscroll=True)

        summarize_text_layout = [
            [summarize_text_menu_label],
            [sg.Text('')],
            [summarize_text_spacer],
            [summarize_text_input_text, summarize_text_input, summarize_text_select_button],
            [sg.Text('')],
            [summarize_text_button],
            [sg.Text('')],
            [summarize_text_output_text],
            [summarize_text_output],
            [sg.Text("", key='-SUMMARIZE-MSG-')]
        ]

        window = sg.Window(tool_name, summarize_text_layout, size=(950, 500), font=("Times New Roman", 15))

        while True:
            event, values = window.read()

            if event == sg.WIN_CLOSED:
                break
            elif event == "Summarize Text":
                input_text_filepath = summarize_text_layout[3][1].get()
                input_text = tool.extract_text(input_text_filepath)
                chunked_input = tool.chunk_text(input_text)
                summarized_text = tool.summarize_text(chunked_input)
                window['-SUMMARIZED-TEXT-'].update(summarized_text)

        window.close()

    while True:
        if window:
            event, _ = window.read()

            if event == sg.WIN_CLOSED:
                break
        else:
            break

    if window:  # Close window if it's open
        window.close()


if __name__ == "__main__":
    main_menu()
