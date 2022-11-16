"""
Set up and settings for the module, including defining constants, setting up logging, and setting up variables.
"""
import sys
import os
import logging

# Define constants
CURRENT_DIR = os.path.dirname(os.path.realpath(__file__))

# Set up logging
def get_logger(name):
    logging.basicConfig(filename=os.path.join(CURRENT_DIR, 'log.log'),
                        filemode='a',
                        encoding='utf-8',
                        format='%(asctime)s %(name)8s %(levelname)-8s %(message)s',
                        datefmt='%Y-%m-%d %H:%M:%S',
                        level=logging.INFO)
    # Handle uncaught exceptions
    def handle_exception(exc_type, exc_value, exc_traceback):
        logging.error("Uncaught exception", exc_info=(exc_type, exc_value, exc_traceback))
    sys.excepthook = handle_exception

    return logging.getLogger(name)

# Set up variables
def init():

    global searching
    searching=False

    global keyword_results
    keyword_results=[]

    global keyword_instances
    keyword_instances={}
