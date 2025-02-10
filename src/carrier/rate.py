import logging
import pandas as pd
from utils.utils import (
    get_input_path_rate,
    read_data_frame,
)

logger = logging.getLogger()


def get_rate():
    return read_data_frame(get_input_path_rate(), sheet="Actual", use_cols="A,B,D,H")


def get_specific_rate_columns(use_cols):
    return read_data_frame(get_input_path_rate(), sheet="Actual", use_cols=use_cols)


def get_bdus_with_rates(bdu_df):
    rate_df = get_rate()
    df_rate = pd.merge(
        bdu_df,
        rate_df[["Sector/Pila origen", "Destino", "Tarifa_Sub", "IVA"]],
        how="left",
        left_on=["Origen", "Destino"],
        right_on=["Sector/Pila origen", "Destino"],
    )
    df_rate = df_rate[
        [
            "Fecha",
            "Matrícula",
            "Origen",
            "Destino",
            "Remito",
            "Pago a",
            "Peso",
            "Tarifa_Sub",
            "IVA",
        ]
    ]
    df_rate["Total"] = df_rate.Peso * df_rate.Tarifa_Sub
    logger.info(
        "The total number of rows after merge BDUs with Rates dataframe is {value}".format(
            value=df_rate.shape[0]
        )
    )
    return df_rate


def get_tree_field_name(source, destination):
    df_rate = get_specific_rate_columns(use_cols="A,B,G")
    df_rate.columns = ["Sector_pila_origen", "Destino", "nombre_detalle"]
    return df_rate[
        (df_rate.Sector_pila_origen == source) & (df_rate.Destino == destination)
    ].nombre_detalle.iloc[0]
