import os
import pathlib
import PySimpleGUI as sg
"""
GUI application to search for keywords in IFRC documents.
"""


"""
Define the window layout
"""
CURRENT_DIR = os.path.dirname(os.path.realpath(__file__))
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
                sg.Image(os.path.join(CURRENT_DIR, 'static/ifrc_logo.png')),
                sg.VSeparator(),
                sg.Text('Keyword Searcher', key='-TITLE-', font = ('OpenSans-Regular', 16), text_color='Black')
            ],
            [sg.Text('Select a folder with documents to search')],
            [sg.Combo(sorted(sg.user_settings_get_entry('-foldernames-', [])), default_value=sg.user_settings_get_entry('-last foldername-', ''), size=(45, 1), key='-FOLDERNAME-')],
            [sg.FolderBrowse(target='-FOLDERNAME-'), sg.B('Clear History')],
            [sg.Text('Enter keywords (one per line)')],
            [sg.Multiline('\n'.join(sg.user_settings_get_entry('-keywords-', [])), size=(45, 5), key='-KEYWORDS-')],
            [sg.Text('Number of words as padding in results'), sg.InputText(10, size=(5, 1), key='-SET WORD PAD-')],
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
                   icon=os.path.join(CURRENT_DIR, 'static/icon.ico'))
window['-SET PAGE-'].bind("<Return>", "_enter")
window['-DOC VIEWER-'].bind('<Enter>', '_hover')
window['-DOC VIEWER-'].bind('<Leave>', '_away')
window['-RESULTS TABLE-'].bind('<Double-Button-1>', '_double_click')
window['-RESULTS TABLE-'].bind("<Return>", "_enter")

"""
Functions
"""
def highlight_document(file, keyword_instances):
    """
    Highlight keywords in a fitz file, looping per page.

    Parameters
    ----------
    file : fitz file (required)

    keyword_instances (required)
        Dictionary mapping pages to lists of fitz instances.
    """
    for pageno in keyword_instances:
        for inst in keyword_instances[pageno]:
            file[pageno].add_highlight_annot(inst)
    return file

# Loop through files in a folder and search for keywords
def loop_files_search_keywords(filepaths, keywords):
    """
    Loop through a list of files (with paths) and search for keywords in each document.

    Convert to PDF if necessary.

    Parameters
    ----------
    filepaths : list (required)
        List of filepaths to search. Keyword searching will be run on each file.

    keywords : list (required)
        The keywords to search for.
    """
    # Get global variables
    global searching
    global keyword_results
    global keyword_instances
    progress_bar = window['progress']
    percent = window['Percent']
    results_summary_text = window['-RESULTS SUMMARY-']

    # Loop through the files in the folder
    keyword_instances = {}
    for i, filepath in enumerate(filepaths):
        filename = os.path.basename(filepath)
        keyword_instances[filename] = {}

        # Break if the search has been stopped
        if not searching:
            break

        # Check the file extension and if it is not a pdf, skip and print a warning
        file_extension = pathlib.Path(filename).suffix
        if file_extension != '.pdf':
            warning_text = window['-SEARCH WARNING-']
            warning_message = f'Skipping file {filename} as it is not a PDF.'
            if warning_text.get():
                warning_message = f'{warning_text.get()}\n{warning_message}'
            warning_text.update(value=warning_message, visible=True)
            window.refresh()
            continue

        # Open the file and extract words with positions
        file = fitz.open(filepath)
        pdf_words = {}
        for page in file:
            pdf_words[page.number] = list(page.get_text("words"))

        # Search for keywords using PyMuPDF and save to keyword_results
        file_results = []
        for pageno, page in enumerate(file):
            keyword_instances[filename][pageno] = []
            for keyword in keywords:
                instances = page.search_for(keyword)
                if instances:
                    keyword_instances[filename][pageno] += instances

                    # Loop through instances to extract information
                    for instance in instances:

                        # Find the word containing the instance start and the word containing the instance end
                        istart = iend = None
                        for j, word in enumerate(pdf_words[page.number]):
                            if (istart is not None) and (iend is not None): break
                            if istart is None:
                                if (word[1] == instance[1]) and (word[0] <= instance[0]) and (word[2] >= instance[0]):
                                    istart = j
                            if iend is None:
                                if (word[3] == instance[3]) and (word[0] <= instance[2]) and (word[2] >= instance[2]):
                                    iend = j

                        # Get words either side of the target keyword
                        try:
                            word_pad = int(values['-SET WORD PAD-'])
                        except Exception as err:
                            word_pad = 10
                            window['-SET WORD PAD-'].update(10)
                        words_before = pdf_words[page.number][(0 if (istart-word_pad)<0 else istart-word_pad):istart]
                        if (len(words_before)<word_pad) and page.number > 0:
                            words_before = pdf_words[page.number-1][len(words_before)-word_pad:]
                        words_after = pdf_words[page.number][iend+1:iend+word_pad+1]
                        if (len(words_after)<word_pad) and (page.number < len(file)-1):
                            words_after += pdf_words[page.number+1][:word_pad-len(words_after)]
                        word_block = words_before+pdf_words[page.number][istart:iend+1]+words_after

                        # Then get the text block to get the spacing back
                        text_block = page.get_textbox((0, word_block[0][1], page.rect.width, word_block[-1][3])).replace('\xa0', ' ').replace('\uf0b7 \n', ' - ')

                        # Add results to be displayed in the table
                        file_results.append([filename, pageno+1, keyword, text_block])

        file.close()

        # Print the results to the table
        if file_results:
            results_summary['keywords'] += len(file_results)
            results_summary['documents'] += 1
            keyword_results += file_results
            results_summary_text.update(value=f'{results_summary["keywords"]} keywords found in {results_summary["documents"]} documents')

        # Update the progress bar
        percent_completed = 100*(i+1)/len(files_to_search)
        percent.update(value=f'{round(percent_completed, 1)} %')
        progress_bar.update_bar(percent_completed)

    # Update the table and set to not searching
    searching = False
    window['-RESULTS TABLE-'].update(keyword_results)
    window['-SEARCH FOR KEYWORDS-'].update('Search')


"""
Create an event loop
"""
# Create the event loop
keyword_results = keyword_instances = None
searching = False
search_folder = search_keywords = None
doc_viewer_hover = False
display_lists = []
view_doc_viewer = False
open_filename = None
temp_dir = None


while True:
    event, values = window.read()
    update_page = False

    if event == sg.WIN_CLOSED:
        searching = False
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
        if searching:
            searching = False
            window['-SEARCH FOR KEYWORDS-'].update('Search')

        # Else begin searching
        else:
            window['-SEARCH ERROR-'].update(value='')
            searching = True
            open_filename = open_page = open_file = None # Refresh to set everything as closed
            if view_doc_viewer:
                view_doc_viewer = False
                window['-DOC VIEWER COLUMN-'].update(visible=view_doc_viewer)
                window.refresh()
            window['-SEARCH WARNING-'].update(visible=False, value='')
            window['-SAVE MESSAGE-'].update(visible=False, value='')
            window['-SEARCH FOR KEYWORDS-'].update('Cancel')
            keyword_results = []
            results_summary = {'keywords': 0, 'documents': 0}
            window['-RESULTS TABLE-'].update([[]])
            window['-RESULTS SUMMARY-'].update(value=f'0 keywords found in 0 documents')
            window['-TEXTBLOCK-'].update(visible=False, value='')

            # Save the keywords, and folder path for next time
            sg.user_settings_set_entry('-keywords-', list(set(keywords)))
            sg.user_settings_set_entry('-foldernames-', list(set(sg.user_settings_get_entry('-foldernames-', []) + [search_folder, ])))
            sg.user_settings_set_entry('-last foldername-', search_folder)

            # Loop through files in the folder and search for keywords
            files_to_search = sorted(os.listdir(search_folder))
            filepaths_to_search = [os.path.join(search_folder, filename) for filename in files_to_search]
            import fitz
            from threading import Thread
            thread = Thread(target=loop_files_search_keywords, args=(filepaths_to_search, keywords))
            thread.start()

    # Export the results or save all documents containing keywords
    elif event=='-EXPORT RESULTS-':
        export_filename = values['-EXPORT RESULTS-']
        if export_filename:
            if keyword_results:
                import csv
                with open(export_filename, 'w', newline='',  encoding='utf-8') as f:
                    writer = csv.writer(f)
                    writer.writerow(results_headers)
                    writer.writerows(keyword_results)
                window['-SAVE MESSAGE-'].update(value='Results saved successfully', visible=True)
    elif event=='-SAVE KEYWORD DOCUMENTS-':
        export_foldername = values['-SAVE KEYWORD DOCUMENTS-']
        if export_foldername:
            if keyword_instances is not None:
                if keyword_instances.keys():
                    # Loop through the documents containing keywords, applying highlighting, and save
                    for document_name in keyword_instances.keys():
                        keyword_document = fitz.open(os.path.join(search_folder, document_name))
                        keyword_document = highlight_document(keyword_instances=keyword_instances[document_name], file=keyword_document)
                        keyword_document.save(os.path.join(export_foldername, document_name))
                    window['-SAVE MESSAGE-'].update(value='Documents saved successfully', visible=True)

    # Display PDFs with keyword when clicked on in table
    elif event=='-RESULTS TABLE-':
        if keyword_results and values[event]:

            # Get the filename, keyword, and page from the selected row
            selected_row = keyword_results[values[event][0]]
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
                open_file = fitz.open(os.path.join(search_folder, selected_filename))
                total_pages = len(open_file)
                open_filename = selected_filename
                display_lists = [None]*total_pages

                # Set the total pages text
                window['-DOCUMENT NAME-'].update(f'{open_filename}')
                window['-TOTAL PAGES-'].update(f'Total pages: {total_pages}')

                # Add highlighting found by previous keyword searching to each page in the file
                open_file = highlight_document(keyword_instances=keyword_instances[selected_filename],
                                               file=open_file)

    # Double clicking a row in the results table should open the file
    elif event in('-RESULTS TABLE-_double_click', '-RESULTS TABLE-_enter'):
        if keyword_results and values['-RESULTS TABLE-']:

            # Get information on the selected row from the table
            selected_row = keyword_results[values['-RESULTS TABLE-'][0]]
            selected_filename = selected_row[0]
            selected_page = selected_row[1]-1

            # Open the document, apply highlighting, and open
            highlighted_doc = highlight_document(keyword_instances=keyword_instances[selected_filename],
                                                 file=fitz.open(os.path.join(search_folder, selected_filename)))
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
