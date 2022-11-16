import os
import pathlib
import PySimpleGUI as sg
import settings
from document_searcher import DocumentSearcher
from document import Document
"""
GUI application to search for keywords in IFRC documents.
"""
# Set up logging
logger = settings.get_logger("base")


"""
Define the window layout
"""
logger.info('Program starting')
sg.change_look_and_feel('Default1')
new_page = 0
image_elem = sg.Image(key='-DOC VIEWER-', expand_x=True, expand_y=True)
goto = sg.InputText(str(new_page + 1), size=(5, 1), key='-SET PAGE-')
results_headers = ['File name', 'Page', 'Keyword', 'Text block']

# Full layout
layout = [
    [
        sg.Column([
            [
                sg.Image(os.path.join(settings.CURRENT_DIR, 'static/ifrc_nsd_logo.png')),
                sg.VSeparator(),
                sg.Text('Keyword Searcher', key='-TITLE-', font = ('OpenSans-Regular', 16), text_color='Black')
            ],
            [sg.Text('Select a folder with documents to search')],
            [sg.Combo(sorted(sg.user_settings_get_entry('-foldernames-', [])), default_value=sg.user_settings_get_entry('-last foldername-', ''), size=(45, 1), key='-FOLDERNAME-')],
            [sg.FolderBrowse(target='-FOLDERNAME-'), sg.B('Clear History')],
            [sg.Text('Enter keywords (one per line)')],
            [sg.Multiline('\n'.join(sg.user_settings_get_entry('-keywords-', [])), size=(45, 5), key='-KEYWORDS-')],
            [sg.Text('Number of words as padding in results'), sg.InputText(10 if not sg.user_settings_get_entry('-LAST WORD PAD-') else sg.user_settings_get_entry('-LAST WORD PAD-'), size=(5, 1), key='-SET WORD PAD-')],
            [sg.Button('Search', key='-SEARCH FOR KEYWORDS-'), sg.Text('', key='-SEARCH ERROR-', text_color='red')],
            [sg.Text('', key='-SEARCH WARNING-', text_color='red', visible=False)],
            [sg.ProgressBar(max_value=100, orientation='h', size=(20, 20), key='progress'),
            sg.Text("0 %", size=(6, 1), key='Percent')],
            [sg.Text("", key='-RESULTS SUMMARY-')],
            [sg.Table(values=[[]],
                      headings=results_headers,
                      auto_size_columns=False,
                      col_widths=(10, 5, 7, 13),
                      key="-RESULTS TABLE-",
                      enable_events=True,
                      justification='left',
                      expand_y=True)],
            [sg.Multiline('', size=(45, 10), visible=False, key='-TEXTBLOCK-')],
            [sg.InputText('', do_not_clear=False, visible=False, key='-EXPORT RESULTS-', enable_events=True),
            sg.FileSaveAs('Save results', target='-EXPORT RESULTS-', file_types=(("CSV Files", "*.csv"),)),
            sg.InputText('', do_not_clear=False, visible=False, key='-SAVE KEYWORD DOCUMENTS-', enable_events=True),
            sg.FolderBrowse('Save all documents containing keywords', target='-SAVE KEYWORD DOCUMENTS-')],
            [sg.Text('', key='-SAVE MESSAGE-', text_color='green')],
        ], expand_y=True, expand_x=False, key='-SEARCH COLUMN-'),
        sg.VSeparator(),
        sg.Column([
            [
                sg.Text('', key='-DOCUMENT NAME-'),
                sg.Button('Prev', key='-PREV PAGE-'),
                sg.Button('Next', key='-NEXT PAGE-'),
                sg.Text('Page:'),
                goto,
                sg.Text('', key='-TOTAL PAGES-'),
            ],
            [image_elem],
        ], key='-DOC VIEWER COLUMN-', visible=False, scrollable=True, vertical_scroll_only=False, size=(920, None), expand_x=True, expand_y=True)
    ]
]
window = sg.Window('IFRC Keyword Searcher',
                   layout,
                   return_keyboard_events=True,
                   finalize=True,
                   resizable=True,
                   icon=os.path.join(settings.CURRENT_DIR, 'static/ifrc_nsd_logo.ico'))
window['-SET PAGE-'].bind("<Return>", "_enter")
window['-DOC VIEWER-'].bind('<Enter>', '_hover')
window['-DOC VIEWER-'].bind('<Leave>', '_away')
window['-RESULTS TABLE-'].bind('<Double-Button-1>', '_double_click')
window['-RESULTS TABLE-'].bind("<Return>", "_enter")

"""
Create an event loop
"""
# Create the event loop
search_folder = search_keywords = None
doc_viewer_hover = False
display_lists = []
view_doc_viewer = False
open_filename = None
temp_dir = None

settings.init()


while True:
    event, values = window.read()
    update_page = False

    if event == sg.WIN_CLOSED:
        settings.searching = False
        break

    # Clear the search documents folder history
    elif event == 'Clear History':
        sg.user_settings_set_entry('-foldernames-', [])
        sg.user_settings_set_entry('-last foldername-', '')
        window['-FOLDERNAME-'].update(values=[], value='')

    # Search for keywords!
    elif event == '-SEARCH FOR KEYWORDS-':
        search_folder = values['-FOLDERNAME-']
        keywords = [word.strip() for word in values['-KEYWORDS-'].split('\n') if word.strip()!='']

        # Check there is a folder and keywords entered
        if not os.path.isdir(search_folder):
            window['-SEARCH ERROR-'].update(value='Please enter a valid search folder')
            continue
        if not search_folder and not keywords:
            window['-SEARCH ERROR-'].update(value='Please enter a search folder and keywords')
            continue
        elif not search_folder:
            window['-SEARCH ERROR-'].update(value='Please enter a search folder')
            continue
        elif not search_folder:
            window['-KEYWORDS-'].update(value='Please enter keywords')
            continue

        # If searching already, then cancel the search
        if settings.searching:
            logger.info("Keyword searching cancelling")
            settings.searching = False
            window['-SEARCH FOR KEYWORDS-'].update('Search')

        # Else begin searching
        else:
            logger.info("Keyword searching starting")
            window['-SEARCH ERROR-'].update(value='')
            settings.searching = True
            open_filename = open_page = open_file = None # Refresh to set everything as closed
            if view_doc_viewer:
                view_doc_viewer = False
                window['-DOC VIEWER COLUMN-'].update(visible=view_doc_viewer)
                window.refresh()
            window['-SEARCH WARNING-'].update(visible=False, value='')
            window['-SAVE MESSAGE-'].update(visible=False, value='')
            window['-SEARCH FOR KEYWORDS-'].update('Cancel')
            window['-RESULTS TABLE-'].update([[]])
            window['-RESULTS SUMMARY-'].update(value=f'0 keywords found in 0 documents')
            window['-TEXTBLOCK-'].update(visible=False, value='')

            # Save the keywords, and folder path for next time
            sg.user_settings_set_entry('-keywords-', list(set(keywords)))
            sg.user_settings_set_entry('-foldernames-', list(set(sg.user_settings_get_entry('-foldernames-', []) + [search_folder, ])))
            sg.user_settings_set_entry('-last foldername-', search_folder)

            # Set the word padding based on the user input
            try:
                word_pad = int(values['-SET WORD PAD-'])
                sg.user_settings_set_entry('-LAST WORD PAD-', word_pad)
            except Exception as err:
                word_pad = 10
                window['-SET WORD PAD-'].update(10)

            # Loop through files in the folder and search for keywords
            files_to_search = sorted(os.listdir(search_folder))
            filepaths_to_search = [os.path.join(search_folder, filename) for filename in files_to_search]
            window['-RESULTS SUMMARY-'].update(value=f'Found {len(filepaths_to_search)} documents to search')
            logger.info(f"Found {len(filepaths_to_search)} files to search")
            import fitz
            from threading import Thread
            thread = Thread(target=DocumentSearcher().search_for_keywords, args=(filepaths_to_search, keywords, word_pad, window))
            thread.start()

    # Export the results or save all documents containing keywords
    elif event=='-EXPORT RESULTS-':
        export_filename = values['-EXPORT RESULTS-']
        if export_filename:
            if settings.keyword_results:
                import csv
                with open(export_filename, 'w', newline='',  encoding='utf-8') as f:
                    writer = csv.writer(f)
                    writer.writerow(results_headers)
                    writer.writerows(settings.keyword_results)
                window['-SAVE MESSAGE-'].update(value='Results saved successfully', visible=True)
    elif event=='-SAVE KEYWORD DOCUMENTS-':
        export_foldername = values['-SAVE KEYWORD DOCUMENTS-']
        if export_foldername:
            if settings.keyword_instances is not None:
                if settings.keyword_instances.keys():
                    # Loop through the documents containing keywords, applying highlighting, and save
                    for document_name in settings.keyword_instances.keys():
                        highlighted_doc = Document(filepath=os.path.join(search_folder, document_name)).highlight_doc(settings.keyword_instances[document_name])
                        highlighted_doc.save(os.path.join(export_foldername, document_name))
                    window['-SAVE MESSAGE-'].update(value='Documents saved successfully', visible=True)

    # Display PDFs with keyword when clicked on in table
    elif event=='-RESULTS TABLE-':
        if settings.keyword_results and values[event]:

            # Get the filename, keyword, and page from the selected row
            selected_row = settings.keyword_results[values[event][0]]
            selected_filename = selected_row[0]
            new_page = selected_row[1]-1
            selected_keyword = selected_row[2]

            # Update the textblock
            window['-TEXTBLOCK-'].update(selected_row[3], visible=True)

            # Clicking to open a NEW file: open the file with fitz and apply highlighting
            if selected_filename!=open_filename:
                update_page = True
                if open_file: open_file.close()

                # Open the file with fitz
                doc = Document(filepath=os.path.join(search_folder, selected_filename))
                total_pages = doc.total_pages
                open_filename = selected_filename
                display_lists = [None]*doc.total_pages

                # Set the total pages text
                window['-DOCUMENT NAME-'].update(f'{open_filename}')
                window['-TOTAL PAGES-'].update(f'Total pages: {doc.total_pages}')

                # Add highlighting found by previous keyword searching to each page in the file
                open_file = doc.highlight_doc(page_rects=settings.keyword_instances[selected_filename])

            # Update the position
            keyword_position = selected_row[4][1]
            page_height = open_file[new_page].rect.height
            window["-DOC VIEWER COLUMN-"].Widget.canvas.yview_moveto(keyword_position/page_height)

    # Double clicking a row in the results table should open the file
    elif event in('-RESULTS TABLE-_double_click', '-RESULTS TABLE-_enter'):
        if settings.keyword_results and values['-RESULTS TABLE-']:

            # Get information on the selected row from the table
            selected_row = settings.keyword_results[values['-RESULTS TABLE-'][0]]
            selected_filename = selected_row[0]
            selected_page = selected_row[1]-1

            # Open the document, apply highlighting, and open
            highlighted_doc = Document(filepath=os.path.join(search_folder, selected_filename)).highlight_doc(page_rects=settings.keyword_instances[selected_filename])
            import tempfile
            if temp_dir is None:
                temp_dir = tempfile.TemporaryDirectory()
            temp_filepath = os.path.join(temp_dir.name, selected_filename)
            highlighted_doc.save(temp_filepath)

            # Try to open the file at the right page
            import webbrowser
            try:
                open_path = pathlib.Path(temp_filepath).as_uri()
                webbrowser.open(f'{open_path}#page={selected_page}') # The page information is being stripped....
            except Exception as err:
                os.startfile(temp_filepath)

    # Change pages of the document
    elif event=='-SET PAGE-_enter':
        try:
            new_page = int(values['-SET PAGE-'])-1
        except:
            pass
    elif event == "-NEXT PAGE-":
        new_page += 1
    elif event == "-PREV PAGE-":
        new_page -= 1
    elif event == "-DOC VIEWER-_hover":
        doc_viewer_hover = True
    elif event == "-DOC VIEWER-_away":
        doc_viewer_hover = False

    # Update the document page if required
    if open_filename is not None:
        if new_page > total_pages-1:
            new_page = total_pages-1
        elif new_page < 0:
            new_page = 0
        if (new_page!=open_page):
            update_page = True

        # Open the page of the document if it has been updated
        if update_page:
            if not view_doc_viewer:
                window['-DOC VIEWER COLUMN-'].update(visible=True)
                view_doc_viewer = True
            if not display_lists[new_page]:  # create if not yet there
                display_lists[new_page] = open_file[new_page].get_displaylist()
            dlist = display_lists[new_page]
            pix = dlist.get_pixmap(alpha=False, matrix=fitz.Matrix(1.5, 1.5))
            image_elem.update(data=pix.tobytes(output='png'))
            open_page = new_page # Set that the currently open page is the new page
            goto.update(str(new_page + 1))

window.close()
