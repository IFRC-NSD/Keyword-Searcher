"""
Document class
"""
import os
import settings
from document import Document

# Set up logging
logger = settings.get_logger("document_searcher")


class DocumentSearcher:
    """
    Module to loop through a given list of documents, search for keywords, and print the output to the window.
    """
    def __init__(self):
        pass


    def search_for_keywords(self, filepaths, search_folder, keywords, word_pad, window):
        """
        Loop through a list of files (with paths) and search for keywords in each document.

        Convert to PDF if necessary.

        Parameters
        ----------
        filepaths : list (required)
            List of filepaths to search. Keyword searching will be run on each file.

        search_folder : str (required)
            Path to the folder which is being searched, which contains the filepaths.
        
        keywords : list (required)
            The keywords to search for.

        word_pad : int (required)
            The number of words to be returned either side of the found keyword.

        window : PySimpleGUI window object (required)
            PySimpleGUI window so that GUI features can be updated as searching is run, such as the progress bar.
        """
        # Get global variables
        settings.keyword_results=None
        settings.keyword_instances={}

        # Loop through the files in the folder
        results_summary = {'keywords': 0, 'documents': 0}
        for i, filepath in enumerate(filepaths):

            # Break if the search has been stopped
            if not settings.searching:
                break

            try:

                # Create the document
                doc = Document(filepath=filepath)
                if doc.file_extension != '.pdf':
                    warning_text = window['-SEARCH WARNING-']
                    warning_message = f'Skipping file {filepath} as it is not a PDF.'
                    if warning_text.get():
                        warning_message = f'{warning_text.get()}\n{warning_message}'
                    warning_text.update(value=warning_message, visible=True)
                    window.refresh()
                    continue

                # Search for keywords in the file
                results = []; instances = {}
                try:
                    results, instances = doc.search_for_keywords(keywords=keywords, word_pad=word_pad)
                    doc.close()
                except Exception as err:
                    logger.exception('Error searching for keywords in document')

                # Update the global keyword results variable. Change the full path to a relative path to display in the results.
                for result in results:
                    result[0] = os.path.relpath(result[0], search_folder)
                if settings.keyword_results is None:
                    settings.keyword_results = results
                else:
                    settings.keyword_results += results
                settings.keyword_instances[os.path.relpath(filepath, search_folder)] = instances

                # Update the progress message
                if results:
                    results_summary['keywords'] += len(results)
                    results_summary['documents'] += 1
                    window['-RESULTS SUMMARY-'].update(value=f'Found {len(filepaths)} documents to search\n{results_summary["keywords"]} keywords found in {results_summary["documents"]} documents')

            # Catch any exceptions and log to the log file
            except Exception as err:
                logger.exception('Error searching keywords in documents')

            # Update the progress bar
            percent_completed = 100*(i+1)/len(filepaths)
            window['Percent'].update(value=f'{round(percent_completed, 1)} %')
            window['progress'].update_bar(percent_completed)

        # Once searching has finished, update the results table and set to not searching
        settings.searching = False
        window['-SEARCH FOR KEYWORDS-'].update('Search')
        window['-RESULTS TABLE-'].update([item[:4] for item in settings.keyword_results])
