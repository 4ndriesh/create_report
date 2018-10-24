import logging

logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s:%(levelname)s:%(message)s')

def exeption_print(err):
    logging.debug(err)