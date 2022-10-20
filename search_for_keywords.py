import os
import re
import time
import pandas as pd
import fitz
from threading import Thread
import PySimpleGUI as sg
"""
GUI application to search for keywords in IFRC documents.
"""


"""
Define the window layout
"""
select_documents_row = [
    sg.Text('Select a folder with documents to search'),
    sg.Combo(sorted(sg.user_settings_get_entry('-foldernames-', [])), default_value=sg.user_settings_get_entry('-last foldername-', ''), size=(50, 1), key='-FOLDERNAME-'),
    sg.FolderBrowse(),
    sg.B('Clear History')
]
keywords_antiwords_row = [
    sg.Text('Enter keywords'),
    sg.Multiline('\n'.join(sg.user_settings_get_entry('-keywords-', [])), size=(30, 5), key='-KEYWORDS-'),
    sg.Text('Enter antiwords'),
    sg.Multiline('\n'.join(sg.user_settings_get_entry('-antiwords-', [])), size=(30, 5), key='-ANTIWORDS-')
]
search_button = [
    sg.Button('Search', key='-SEARCH FOR KEYWORDS-')
]
progress_row = [
    sg.ProgressBar(max_value=100, orientation='h', size=(20, 20), key='progress'),
    sg.Text("0 %", size=(4, 1), key='Percent')
]
results_summary_row = [
    sg.Text("", key='-RESULTS SUMMARY-'),
]
results_table_row = [
    sg.Table(values=[[]],
             headings=['ID', 'File', 'Keyword', 'Page'],
             size=(100, 20),
             key="-RESULTS TABLE-",
             expand_y=True,
             enable_events=True)
]
export_row = [
    sg.InputText('', do_not_clear=False, visible=False, key='-EXPORT RESULTS-', enable_events=True),
    sg.FileSaveAs('Save results', file_types=(("CSV Files", "*.csv"),))
]
cur_page = 0
image_elem = sg.Image(key='-DOC VIEWER-')
goto = sg.InputText(str(cur_page + 1), size=(5, 1), key='-SET PAGE-')

# Full layout
layout = [[
    [select_documents_row, keywords_antiwords_row, search_button, progress_row],
    sg.Column([results_summary_row, results_table_row, export_row]),
    sg.VSeparator(),
    sg.Column([
        [
            sg.Button('Prev', key='-PREV PAGE-'),
            sg.Button('Next', key='-NEXT PAGE-'),
            sg.Text('Page:'),
            goto,
        ],
        [image_elem],
    ])
]]
window = sg.Window('IFRC Keyword Searcher', layout, finalize=True)
window['-SET PAGE-'].bind("<Return>", "_enter")
window['-DOC VIEWER-'].bind('<Enter>', '_hover')

"""
Functions
"""
# Loop through files in a folder and search for keywords
def loop_files_search_keywords(folderpath, keywords):
    global searching
    global keyword_results
    global display_lists
    progress_bar = window['progress']
    percent = window['Percent']
    results_summary_text = window['-RESULTS SUMMARY-']

    # Loop through the files in the folder
    files_to_search = sorted(os.listdir(folderpath))
    keyword_instances = []
    for i, filename in enumerate(files_to_search):

        # Break if the search has been stopped
        if not searching:
            break

        # Search for keywords using PyMuPDF
        file_results = []
        file = fitz.open(os.path.join(folderpath, filename))
        display_lists = [None]*len(file)
        for keyword in keywords:
            for pageno, page in enumerate(file):
                page_instances = page.search_for(keyword)
                file_results += [[id, filename, keyword, pageno] for id in range(len(keyword_instances), len(keyword_instances+page_instances))]
                keyword_instances += page_instances
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

    # Update the table
    window['-RESULTS TABLE-'].update(keyword_results)
    window['-SEARCH FOR KEYWORDS-'].update('Search')


"""
Create an event loop
"""
# Create the event loop
searching = False
open_filename = None
display_lists = []

while True:
    event, values = window.read()
    force_page = False

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
            window['-SEARCH FOR KEYWORDS-'].update('Cancel')
            keyword_results = []
            results_summary = {'keywords': 0, 'documents': 0}
            window['-RESULTS TABLE-'].update([[]])
            window['-RESULTS SUMMARY-'].update(value=f'0 keywords found in 0 documents')

            # Save the keywords, antiwords, and folder path for next time
            keywords = [word.strip() for word in values['-KEYWORDS-'].split('\n') if word.strip()!='']
            sg.user_settings_set_entry('-keywords-', list(set(keywords)))
            antiwords = [word.strip() for word in values['-ANTIWORDS-'].split('\n') if word.strip()!='']
            sg.user_settings_set_entry('-antiwords-', list(set(antiwords)))
            sg.user_settings_set_entry('-foldernames-', list(set(sg.user_settings_get_entry('-foldernames-', []) + [values['-FOLDERNAME-'], ])))
            sg.user_settings_set_entry('-last foldername-', values['-FOLDERNAME-'])

            # Loop through files in the folder and search for keywords
            thread = Thread(target=loop_files_search_keywords, args=(values['-FOLDERNAME-'], keywords))
            thread.start()

    # Export the results
    elif event=='-EXPORT RESULTS-':
        export_filename = values['-EXPORT RESULTS-']
        if export_filename:
            keyword_results.to_csv(export_filename, index=False)

    # Display PDFs with keyword when clicked on in table
    elif event=='-RESULTS TABLE-':
        if keyword_results:

            # Get information on the selected row
            selected_filename = keyword_results[values[event][0]][1]
            if selected_filename!=open_filename: # We click to open a new file
                open_file = fitz.open(os.path.join(values['-FOLDERNAME-'], selected_filename))
                open_filename = selected_filename

            # Open the document, using the saved one if already got
            cur_page = open_page = 0
            update_page = True

    # Change pages of the document
    elif event=='-SET PAGE-'+'_enter':
        try:
            cur_page = int(values['-SET PAGE-'])-1
        except:
            pass
    elif event in ("-NEXT PAGE-",):
        cur_page += 1
    elif event in ("-PREV PAGE-",):
        cur_page -= 1

    # Update the document page if required
    if open_filename is not None:
        if cur_page > len(open_file)-1:
            cur_page = len(open_file)-1
        elif cur_page < 0:
            cur_page = 0
        if (cur_page!=open_page):
            update_page = True
        if update_page:
            if not display_lists[cur_page]:  # create if not yet there
                display_lists[cur_page] = open_file[cur_page].get_displaylist()
            dlist = display_lists[cur_page]
            pix = dlist.get_pixmap(alpha=False)
            image_elem.update(data=pix.tobytes(output='png'))
            goto.update(str(cur_page + 1))


window.close()
