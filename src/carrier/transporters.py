import logging
import pandas as pd
from utils.utils import (
    get_input_path_file_xslx,
    get_dinama_file_name,
    get_dinama_sheet_name,
    read_data_frame,
)

logger = logging.getLogger()


def get_truck_registration():
    path_file_dinama = get_input_path_file_xslx() + get_dinama_file_name()
    sheet_name = get_dinama_sheet_name()
    df_truck_id = read_data_frame(path_file_dinama, sheet=sheet_name, use_cols="A,B,C")
    df_truck_id["MATRICULA"] = df_truck_id["MATRICULA"].str.replace(" ", "")
    df_truck_id.columns = ["Matrícula", "Transportista", "Rut"]
    df_truck_id = df_truck_id[["Matrícula", "Transportista"]]
    return df_truck_id


def get_truck_id_with_rates(df_bdu_with_rates):
    df_truck_id = get_truck_registration()
    df_truck_id_with_rates = pd.merge(
        df_truck_id,
        df_bdu_with_rates,
        how="right",
        left_on=["Matrícula"],
        right_on=["Matrícula"],
    )
    df_truck_id_with_rates = df_truck_id_with_rates.sort_values(by=["Remito"])
    df_truck_id_with_rates = df_truck_id_with_rates[
        [
            "Fecha",
            "Matrícula",
            "Peso",
            "Tarifa_Sub",
            "Total",
            "Origen",
            "Destino",
            "Remito",
            "Pago a",
            "Transportista",
        ]
    ]
    return df_truck_id_with_rates


def get_transporters_rut():
    path_file_dinama = get_input_path_file_xslx() + get_dinama_file_name()
    sheet_name = get_dinama_sheet_name()
    df_transporters_rut = read_data_frame(
        path_file_dinama, sheet=sheet_name, use_cols="B,C,D,E,F,G"
    )
    df_transporters_rut.columns = [
        "Transportista",
        "Rut",
        "BANCO",
        "TIPO CUENTA",
        "CUENTA",
        "ANTERIOR",
    ]
    df_transporters_rut.drop_duplicates(keep="first", inplace=True)
    return df_transporters_rut
