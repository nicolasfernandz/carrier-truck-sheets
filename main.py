import json
import logging
import os
import datetime
from utils import convert_unicode_dict
from json_minify import json_minify

from carrierfreight import CarrierFreight

#Aplication logger
log_filename = 'log_' + '.log'
logger = logging.getLogger()
logging.basicConfig(filename=log_filename,level=logging.DEBUG)

def main():
    """ Main script to run the Program"""
    # current directory
    directory = os.path.dirname(os.path.abspath(__file__))
    configuration_path = getProperties()

    # Get application logger
    logger.info("Start travel processing: {}".format(datetime.datetime.now()))

    with open(configuration_path) as f:
        json_string = f.read()
        main_graph = json.loads(json_minify(json_string))
        main_graph_str = convert_unicode_dict(main_graph)

        logger.info("the configuration properties path has been read correctly: {}".format(configuration_path))

        input_paths = main_graph_str["input"]
        type = main_graph_str["type"]

        if(type == "carrierfreight"):
            CarrierFreight(main_graph_str=main_graph_str, input_paths=input_paths).run()
        elif(type == 'control'):
            print('Not implemented yet.')

        logger.info("The process has been completed")

def getProperties():
    """ read from file.properties the main files path"""
    file = open('file.properties')
    lines = file.readlines()
    return lines[0]

if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        logger.error('Failed to execute program: {}'.format(str(e), exc_info=True))
