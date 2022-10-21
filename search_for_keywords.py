import os
import re
import time
import shutil
import pathlib
from threading import Thread
import pandas as pd
import fitz
import PySimpleGUI as sg
"""
GUI application to search for keywords in IFRC documents.
"""


"""
Define the window layout
"""
sg.change_look_and_feel('Default1')
new_page = 0
image_elem = sg.Image(key='-DOC VIEWER-', expand_x=True, expand_y=True)
goto = sg.InputText(str(new_page + 1), size=(5, 1), key='-SET PAGE-')
results_headers = ['File', 'Page', 'Keyword', 'Count']

# Full layout
layout = [
    [
        sg.Column([
            [
                sg.Image('./static/ifrc_logo_small.png'),
                sg.VSeparator(),
                sg.Text('Keyword Searcher', key='-TITLE-', font = ('OpenSans-Regular', 16), text_color='Black')
            ],
            [sg.Text('Select a folder with documents to search')],
            [sg.Combo(sorted(sg.user_settings_get_entry('-foldernames-', [])), default_value=sg.user_settings_get_entry('-last foldername-', ''), size=(35, 1), key='-FOLDERNAME-')],
            [sg.FolderBrowse(target='-FOLDERNAME-'), sg.B('Clear History')],
            [sg.Text('Enter keywords (one per line)')],
            [sg.Multiline('\n'.join(sg.user_settings_get_entry('-keywords-', [])), size=(35, 5), key='-KEYWORDS-')],
            [sg.Button('Search', key='-SEARCH FOR KEYWORDS-')],
            [sg.Text('', key='-SEARCH WARNING-', text_color='red', visible=False)],
            [sg.ProgressBar(max_value=100, orientation='h', size=(20, 20), key='progress'),
            sg.Text("0 %", size=(4, 1), key='Percent')],
            [sg.Text("", key='-RESULTS SUMMARY-')],
            [sg.Table(values=[[]],
                      headings=results_headers,
                      auto_size_columns=False,
                      max_col_width=10,
                      def_col_width=10,
                      key="-RESULTS TABLE-",
                      enable_events=True,
                      expand_y=True)],
            [sg.InputText('', do_not_clear=False, visible=False, key='-EXPORT RESULTS-', enable_events=True),
            sg.FileSaveAs('Save results', target='-EXPORT RESULTS-', file_types=(("CSV Files", "*.csv"),)),
            sg.InputText('', do_not_clear=False, visible=False, key='-SAVE KEYWORD DOCUMENTS-', enable_events=True),
            sg.FolderBrowse('Save all documents containing keywords', target='-SAVE KEYWORD DOCUMENTS-')],
            [sg.Text('', key='-SAVE MESSAGE-', text_color='green')],
        ], expand_y=True),
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
        ], key='-DOC VIEWER COLUMN-', visible=False, expand_x=True, expand_y=True)
    ]
]
window = sg.Window('IFRC Keyword Searcher',
                   layout,
                   return_keyboard_events=True,
                   finalize=True,
                   resizable=True)
window['-SET PAGE-'].bind("<Return>", "_enter")
window['-DOC VIEWER-'].bind('<Enter>', '_hover')
window['-DOC VIEWER-'].bind('<Leave>', '_away')

"""
Functions
"""
# Loop through files in a folder and search for keywords
def loop_files_search_keywords(filepaths, keywords):
    """
    Loop through a list of files (with paths) and search for keywords in each document.^M

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

        # Search for keywords using PyMuPDF
        file = fitz.open(filepath)
        file_results = []
        for pageno, page in enumerate(file):
            keyword_instances[filename][pageno] = []
            page_keywords = []
            page_instances = 0
            for keyword in keywords:
                instances = page.search_for(keyword)
                if instances:
                    page_keywords.append(keyword)
                    page_instances += len(instances)
                    keyword_instances[filename][pageno] += instances
            if page_keywords:
                file_results.append([filename, pageno+1, ', '.join(page_keywords), page_instances])
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
searching = False
doc_viewer_hover = False
display_lists = []
view_doc_viewer = False
open_filename = None

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
        if not values['-FOLDERNAME-'] or not values['-KEYWORDS-']:
            continue

        # If searching already, then cancel the search
        if searching:
            searching = False
            window['-SEARCH FOR KEYWORDS-'].update('Search')

        # Else begin searching
        else:
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

            # Save the keywords, and folder path for next time
            keywords = [word.strip() for word in values['-KEYWORDS-'].split('\n') if word.strip()!='']
            sg.user_settings_set_entry('-keywords-', list(set(keywords)))
            sg.user_settings_set_entry('-foldernames-', list(set(sg.user_settings_get_entry('-foldernames-', []) + [values['-FOLDERNAME-'], ])))
            sg.user_settings_set_entry('-last foldername-', values['-FOLDERNAME-'])

            # Loop through files in the folder and search for keywords
            files_to_search = sorted(os.listdir(values['-FOLDERNAME-']))
            filepaths_to_search = [os.path.join(values['-FOLDERNAME-'], filename) for filename in files_to_search]
            thread = Thread(target=loop_files_search_keywords, args=(filepaths_to_search, keywords))
            thread.start()

    # Export the results or save all documents containing keywords
    elif event=='-EXPORT RESULTS-':
        export_filename = values['-EXPORT RESULTS-']
        if export_filename:
            results_df = pd.DataFrame(data=keyword_results, columns=results_headers)
            results_df.to_csv(export_filename, index=False)
            window['-SAVE MESSAGE-'].update(value='Results saved successfully', visible=True)
    elif event=='-SAVE KEYWORD DOCUMENTS-':
        export_foldername = values['-SAVE KEYWORD DOCUMENTS-']
        if export_foldername:
            for keyword_document in keyword_instances.keys():
                shutil.copyfile(os.path.join(values['-FOLDERNAME-'], keyword_document),
                                os.path.join(export_foldername, keyword_document))
            window['-SAVE MESSAGE-'].update(value='Documents saved successfully', visible=True)

    # Display PDFs with keyword when clicked on in table
    elif event=='-RESULTS TABLE-':
        if keyword_results and values[event]:

            # Get the filename, keyword, and page from the selected row
            selected_row = keyword_results[values[event][0]]
            selected_filename = selected_row[0]
            new_page = selected_row[1]-1
            selected_keyword = selected_row[2]

            # Clicking to open a NEW file: open the file with fitz and apply highlighting
            if selected_filename!=open_filename:
                update_page = True
                if open_file: open_file.close()

                # Open the file with fitz
                open_file = fitz.open(os.path.join(values['-FOLDERNAME-'], selected_filename))
                total_pages = len(open_file)
                open_filename = selected_filename
                display_lists = [None]*total_pages

                # Set the total pages text
                window['-DOCUMENT NAME-'].update(f'{open_filename}')
                window['-TOTAL PAGES-'].update(f'Total pages: {total_pages}')

                # Add highlighting found by previous keyword searching to each page in the file
                for pageno in keyword_instances[selected_filename]:
                    for inst in keyword_instances[selected_filename][pageno]:
                        open_file[pageno].add_highlight_annot(inst)

    # Change pages of the document
    elif event=='-SET PAGE-'+'_enter':
        try:
            new_page = int(values['-SET PAGE-'])-1
        except:
            pass
    elif event in ("-NEXT PAGE-",):
        new_page += 1
    elif event in ("-PREV PAGE-",):
        new_page -= 1
    elif event == "-DOC VIEWER-_hover":
        doc_viewer_hover = True
    elif event == "-DOC VIEWER-_away":
        doc_viewer_hover = False
    elif (event == "MouseWheel:Down") and doc_viewer_hover:
        new_page += 1
    elif (event == "MouseWheel:Up") and doc_viewer_hover:
        new_page -= 1

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
            pix = dlist.get_pixmap(alpha=False)
            image_elem.update(data=pix.tobytes(output='png'))
            open_page = new_page # Set that the currently open page is the new page
            goto.update(str(new_page + 1))


window.close()
