import os
import re
import time
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
image_elem = sg.Image(key='-DOC VIEWER-')
goto = sg.InputText(str(new_page + 1), size=(5, 1), key='-SET PAGE-')

# Full layout
layout = [
    [
        sg.Image('./static/ifrc_logo_small.png'),
        sg.VSeparator(),
        sg.Text('Keyword Searcher', key='-TITLE-', font = ('OpenSans-Regular', 16), text_color='Black')
    ],
    [
        sg.Column([
            [sg.Text('Select a folder with documents to search')],
            [sg.Combo(sorted(sg.user_settings_get_entry('-foldernames-', [])), default_value=sg.user_settings_get_entry('-last foldername-', ''), size=(35, 1), key='-FOLDERNAME-')],
            [sg.FolderBrowse(), sg.B('Clear History')],
            [sg.Text('Enter keywords')],
            [sg.Multiline('\n'.join(sg.user_settings_get_entry('-keywords-', [])), size=(35, 5), key='-KEYWORDS-')],
            [sg.Button('Search', key='-SEARCH FOR KEYWORDS-')],
            [sg.ProgressBar(max_value=100, orientation='h', size=(20, 20), key='progress'),
            sg.Text("0 %", size=(4, 1), key='Percent')],
            [sg.Text("", key='-RESULTS SUMMARY-')],
            [sg.Table(values=[[]],
                      headings=['File', 'Page', 'Keyword', 'Count'],
                      size=(35, 35),
                      auto_size_columns=False,
                      max_col_width=10,
                      def_col_width=10,
                      key="-RESULTS TABLE-",
                      enable_events=True)],
            [sg.InputText('', do_not_clear=False, visible=False, key='-EXPORT RESULTS-', enable_events=True),
            sg.FileSaveAs('Save results', file_types=(("CSV Files", "*.csv"),))]
        ]),
        sg.VSeparator(),
        sg.Column([
            [
                sg.Button('Prev', key='-PREV PAGE-'),
                sg.Button('Next', key='-NEXT PAGE-'),
                sg.Text('Page:'),
                goto,
                sg.Text('', key='-TOTAL PAGES-'),
            ],
            [image_elem],
        ], key='-DOC VIEWER COLUMN-', visible=False)
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
def loop_files_search_keywords(folderpath, keywords):
    global searching
    global keyword_results
    global keyword_instances
    progress_bar = window['progress']
    percent = window['Percent']
    results_summary_text = window['-RESULTS SUMMARY-']

    # Loop through the files in the folder
    files_to_search = sorted(os.listdir(folderpath))
    keyword_instances = {}
    for i, filename in enumerate(files_to_search):
        keyword_instances[filename] = {}

        # Break if the search has been stopped
        if not searching:
            break

        # Search for keywords using PyMuPDF
        file_results = []
        file = fitz.open(os.path.join(folderpath, filename))
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

    # Update the table
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
            open_filename = open_page = open_file = None # Refresh to set everything as closed
            if view_doc_viewer:
                view_doc_viewer = False
                window['-DOC VIEWER COLUMN-'].update(visible=view_doc_viewer)
                window.refresh()
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
            thread = Thread(target=loop_files_search_keywords, args=(values['-FOLDERNAME-'], keywords))
            thread.start()

    # Export the results
    elif event=='-EXPORT RESULTS-':
        export_filename = values['-EXPORT RESULTS-']
        if export_filename:
            keyword_results.to_csv(export_filename, index=False)

    # Display PDFs with keyword when clicked on in table
    elif event=='-RESULTS TABLE-':
        if keyword_results and values[event]:

            # Get the filename, keyword, and page from the selected row
            selected_row = keyword_results[values[event][0]]
            selected_filename = selected_row[0]
            new_page = selected_row[1]-1
            selected_keyword = selected_row[2]

            # Clicking to open a new file
            if selected_filename!=open_filename:
                update_page = True
                if open_file:
                    open_file.close()
                open_file = fitz.open(os.path.join(values['-FOLDERNAME-'], selected_filename))
                open_filename = selected_filename
                display_lists = [None]*len(open_file)

                # Set the total pages text
                window['-TOTAL PAGES-'].update(f'Total pages: {len(open_file)}')

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
        if new_page > len(open_file)-1:
            new_page = len(open_file)-1
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
