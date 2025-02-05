import json
import logging
from json_minify import json_minify
import six
import pandas as pd
import os

logger = logging.getLogger()


def is_NaN(num):
    return num != num


def get_propertie_file():
    """read from file.properties the main files path"""
    file = open("./file.properties")
    lines = file.readlines()
    return lines[0]


def get_properties():
    configuration_path = get_propertie_file()
    with open(configuration_path) as f:
        json_string = f.read()
        main_graph = json.loads(json_minify(json_string))
        main_graph_str = convert_unicode_dict(main_graph)

        logger.info(
            "the configuration properties path has been read correctly: {}".format(
                configuration_path
            )
        )
    return main_graph_str


def get_contractor():
    main_graph_str = get_properties()
    return main_graph_str["contractor"]


def get_input_paths():
    main_graph_str = get_properties()
    return main_graph_str["input"]


def validate_input_properties():
    main_graph_str = get_properties()
    input_paths = main_graph_str["input"]

    if not isinstance(main_graph_str, dict):
        raise Exception("'main_graph_str' must be a dictionary")

    if not isinstance(input_paths, dict):
        raise Exception("'input_paths' must be a dictionary")


def get_input_path_file_xslx():
    input_paths = get_input_paths()
    return input_paths["input_path_file_xslx"] + input_paths["input_file_directory"]


def get_output_path():
    input_paths = get_input_paths()
    return input_paths["output_path"] + input_paths["input_file_directory"]


def get_month_directory():
    input_paths = get_input_paths()
    return input_paths["input_file_directory"]


def get_input_path_rate():
    input_paths = get_input_paths()
    return input_paths["input_path_rates"]


def get_dinama_file_name():
    input_paths = get_input_paths()
    return input_paths["dinama"]


def get_dinama_sheet_name():
    input_paths = get_input_paths()
    return input_paths["dinama_sheet_name"]


def get_fuel_file_name():
    input_paths = get_input_paths()
    return input_paths["fuel"]


def get_process_month():
    input_paths = get_input_paths()
    return input_paths["month"]


def get_sheet_list_names():
    main_graph_str = get_properties()
    bdus = main_graph_str["bdus"]
    return bdus["sheets_names"]


def get_sheets_info_dict():
    main_graph_str = get_properties()
    return main_graph_str["sheets"]


def convert_unicode_dict(input_val):
    """Convert
    :param input_val: input dictionary
    :return: dictionary
    """
    if isinstance(input_val, dict):
        return {
            convert_unicode_dict(key): convert_unicode_dict(value)
            for key, value in input_val.items()
        }
    elif isinstance(input_val, list):
        return [convert_unicode_dict(element) for element in input_val]
    elif isinstance(input_val, six.string_types):
        return str(input_val)
    else:
        return input_val


def is_fuel_present():
    return os.path.isfile(get_input_path_file_xslx() + get_fuel_file_name())


def read_data_frame(path, sheet=None, use_cols=None):
    if use_cols:
        return pd.read_excel(path, sheet_name=sheet, usecols=use_cols)
    else:
        return pd.read_excel(path, sheet_name=sheet)
