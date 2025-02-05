import logging
import pandas as pd
from utils.utils import (
    get_input_path_file_xslx,
    get_fuel_file_name,
    is_fuel_present,
    read_data_frame,
)

logger = logging.getLogger()


def get_fuel():
    if is_fuel_present():
        logger.info(
            "processing fuel form {fuel_sheet_name}".format(
                fuel_sheet_name=get_fuel_file_name()
            )
        )
        df_fuel_keys = read_data_frame(
            get_input_path_file_xslx() + get_fuel_file_name(), sheet=None
        )
        sheet_name_list = []
        for sheet_name in df_fuel_keys.keys():
            sheet_name_list.append(sheet_name)

        df_fuel_concat = pd.DataFrame(
            columns=[
                "Fecha de Emisión",
                "Matrícula",
                "Comprobante",
                "Debe",
                "Transportista",
                "Estacion",
            ]
        )
        for sheet_name in sheet_name_list:
            df_fuel_read = read_data_frame(
                get_input_path_file_xslx() + get_fuel_file_name(),
                sheet=sheet_name,
                use_cols="A,B,C,D,E,F",
            )
            df_fuel_read = df_fuel_read[
                [
                    "Fecha de Emisión",
                    "Matrícula",
                    "Comprobante",
                    "Debe",
                    "Transportista",
                    "Estacion",
                ]
            ]
            df_fuel_read["Matrícula"].fillna("", inplace=True)
            df_fuel_concat = pd.concat(
                [df_fuel_concat, df_fuel_read], ignore_index=True
            )
            logger.info(
                "{value} rows have been added to the fuel dataframe from {carrier}".format(
                    value=df_fuel_read.shape[0], carrier=sheet_name
                )
            )
        df_fuel = df_fuel_concat
        logger.info(
            "total number of rows of the fuel dataframe is {value}".format(
                value=df_fuel.shape[0]
            )
        )
    df_fuel.columns = ['Fecha de Emisión', 'Camión', 'Comprobante', 'Importe', 'Transportista', 'Estacion']
    return df_fuel
