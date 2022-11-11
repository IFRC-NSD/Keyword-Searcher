import os
import pathlib
import fitz

class Document:
    """
    Class to handle documents and in particular search for keywords.

    Parameters
    ----------
    filepath : str (required)
        Path to the document.
    """
    def __init__(self, filepath):
        self.filename = os.path.basename(filepath)
        self.file_extension = pathlib.Path(self.filename).suffix
        self.doc = fitz.open(filepath)


    def search_for_keywords(self, keywords, word_pad=10):
        """
        Search for keywords in the document.

        Parameters
        ----------
        keywords : list (required)
            List of keywords or key phrases to search for.

        word_pad : int (default=10)
            Number of words to return either side of the keywords found.

        Returns
        -------
        doc_results : list
            List of lists, where each list is a row of a dataset, with information on the filename, page number, keyword, and text block.

        doc_instances : dict
            Contains the list of keyword search instances found for each page of the document.
        """
        # Get the document words
        self.words = self.get_words()

        # Loop through the pages of the document and search for keywords in each page
        doc_results = []
        doc_instances = {}
        for pageno, page in enumerate(self.doc):
            doc_instances[pageno] = []
            for keyword in keywords:

                # Get the keyword instances
                instances = page.search_for(keyword)
                if instances:
                    doc_instances[pageno] += instances
                    for instance in instances:

                        # Get a nuber of words (word_pad) either side of the keyword/ phrase, and get the bounding rect
                        word_start, word_end = self.find_bounding_words(words=self.words[pageno], rect=instance)
                        word_count_left, first_word = self.iterate_words_limit(words=reversed(self.words[pageno][:word_start]),
                                                                               limit=word_pad)
                        word_count_right, last_word = self.iterate_words_limit(words=self.words[pageno][word_end+1:],
                                                                               limit=word_pad)
                        text_block = page.get_textbox((0,
                                                       self.words[pageno][word_start][1] if first_word is None else first_word[1],
                                                       page.rect.width,
                                                       self.words[pageno][word_end][3] if last_word is None else last_word[3]))

                        # Get overflow words on the previous and next page
                        if (word_count_left<word_pad) and pageno > 0:
                            word_count_prev_page, first_word_prev_page = self.iterate_words_limit(words=reversed(self.words[pageno-1]),
                                                                                                  limit=word_pad-word_count_left)
                            text_block_prev_page = self.doc[pageno-1].get_textbox((0, first_word_prev_page[1], self.doc[pageno-1].rect.width, self.doc[pageno-1].rect.height))
                            text_block = text_block_prev_page + '\n\n' + text_block
                        if (word_count_right<word_pad) and (pageno < len(self.doc)-1):
                            word_count_next_page, last_word_next_page = self.iterate_words_limit(words=self.words[pageno+1],
                                                                                                 limit=word_pad-word_count_right)
                            text_block_next_page = self.doc[pageno+1].get_textbox((0, 0, self.doc[pageno-1].rect.width, last_word_next_page[3]))
                            text_block = text_block + '\n\n' + text_block_next_page

                        # Remove funny characters
                        text_block = self.tidy_text(text_block)

                        # Add results to be displayed in the table
                        doc_results.append([self.filename, pageno+1, keyword, text_block])

        return doc_results, doc_instances


    def iterate_words_limit(self, words, limit):
        """
        Iterate through words, counting numbers of words accounting for special space characters, and stop when the limit is reached.

        Parameters
        ----------
        words : list of fitz word objects (required)
            List of fitz words to iterate through, where the fifth element of each word is the word as a string.

        limit : int (required)
            Stop after this many words is reached.

        Returns
        -------
        count : int
            Number of words counted (may be greater than limit due to special characters).

        word : fitz word object
            The last word reached, where the first four elements are the rect and the fifth is the word as a string.
        """
        count = 0; word = None
        for word in words:
            count += len(word[4].replace('\xa0', ' ').strip().split())
            if count >= limit:
                break
        return count, word


    def find_bounding_words(self, words, rect):
        """
        Find the words in the document which contain the given rect.
        Find the word which contains the start of the rect, and the word which contains the end of the rect.
        Return the index of the left word and the right word.

        Parameters
        ----------
        words : list of fitz word objects (required)
            List of fitz words to look through, where the first four elements are the coordinates of the word.

        rect : list or tuple (required)
            The coordinates to look for, where the first two elements are the top left corner (x1, y1) and the second two are the top right (x2, y2).

        Returns
        -------
        istart : int
            Index of the word containing the left of the rectangle.

        iend : int
            Index of the word containing the right of the rectangle.
        """
        istart = iend = None
        for i, word in enumerate(words):
            if (istart is not None) and (iend is not None): break
            if istart is None:
                if (word[1] == rect[1]) and (word[0] <= rect[0]) and (word[2] >= rect[0]):
                    istart = i
            if iend is None:
                if (word[3] == rect[3]) and (word[0] <= rect[2]) and (word[2] >= rect[2]):
                    iend = i

        return istart, iend


    def get_words(self):
        """
        Extract words from the document in order.
        Remove page numbers.

        Returns
        -------
        doc_words : dict
            List of fitz word objects for each page of the document.
        """
        doc_words = {}
        for page in self.doc:
            page_words = sorted(list(page.get_text("words")), key=lambda word: [word[1], word[0]])
            # Remove page numbers: if the last word is a long way below the previous word and an integer
            try:
                if (page_words[-1][1]-page_words[-2][3]) > 2*(page_words[-2][3]-page_words[-2][1]):
                    if int(page_words[-1][4].strip()):
                        page_words = page_words[:-1]
            except Exception as err:
                continue
            doc_words[page.number] = page_words

        return doc_words


    def tidy_text(self, txt):
        """
        Tidy a block of text, removing and replacing special characters.

        Parameters
        ----------
        txt : string (required)
            The text to be tidied.
        """
        txt = txt.replace('\xa0', ' ')\
                  .replace('\uf0b7 \n', ' - ')\
                  .strip()
        return txt


    def close(self):
        """
        Close the open document.
        """
        self.doc.close()
