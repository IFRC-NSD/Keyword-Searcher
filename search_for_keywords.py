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
             headings=['File', 'Keyword', 'Text extract'],
             size=(100, 20),
             key="-RESULTS TABLE-",
             expand_y=True,
             enable_events=True)
]
export_row = [
    sg.InputText('', do_not_clear=False, visible=False, key='-EXPORT RESULTS-', enable_events=True),
    sg.FileSaveAs('Save results', file_types=(("CSV Files", "*.csv"),))
]
# Full layout
layout = [[
    [select_documents_row, keywords_antiwords_row, search_button, progress_row],
    sg.Column([results_summary_row, results_table_row, export_row]),
    sg.VSeparator(),
    sg.Column([
        [sg.Text("", key='-RESULT FILENAME-')],
        [sg.Text("", key='-RESULT TEXT-', size=(100,20))]
    ])
]]

window = sg.Window('IFRC Keyword Searcher', layout)

"""
Functions
"""
# Loop through files in a folder and search for keywords
def loop_files_search_keywords(folderpath):
    global searching
    global keyword_results
    progress_bar = window['progress']
    percent = window['Percent']
    results_summary_text = window['-RESULTS SUMMARY-']

    # Loop through the files in the folder
    files_to_search = sorted(os.listdir(folderpath))
    for i, filename in enumerate(files_to_search):

        # Break if the search has been stopped
        if not searching:
            break

        # Extract the document content and search for keywords
        file_path = os.path.join(folderpath, filename)
        #file_content = parser.from_file(file_path)['content']

        results = [[filename, 'alex', 'alex alex alex']] # GET THE KEYWORD RESULTS

        # Print the results to the table
        if results:
            results_summary['keywords'] += len(results)
            results_summary['documents'] += 1
            keyword_results += results
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
while True:
    event, values = window.read()

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

        # If searching, then cancel the search
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
            thread = Thread(target=loop_files_search_keywords, args=(values['-FOLDERNAME-'],))
            thread.start()

    # Display PDFs with keyword when clicked on in table
    elif event=='-RESULTS TABLE-':
        if keyword_results:

            # Get information on the selected row
            selected_filename = keyword_results[values[event][0]]
            selected_text = keyword_results[values[event][0]]
            #selected_filepath = os.path.join(values['-FOLDERNAME-'], selected_filename)

            # Display the pdf at the correct position
            window['-RESULT FILENAME-'].update(selected_filename)
            window['-RESULT TEXT-'].update(selected_text)

    # Export the results
    elif event=='-EXPORT RESULTS-':
        export_filename = values['-EXPORT RESULTS-']
        if export_filename:
            keyword_results.to_csv(export_filename, index=False)


window.close()
