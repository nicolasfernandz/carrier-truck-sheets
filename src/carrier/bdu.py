import logging
import pandas as pd
from utils.utils import (
    get_input_path_file_xslx,
    get_sheet_list_names,
    get_sheets_info_dict,
    read_data_frame,
)

logger = logging.getLogger()


def get_all_bdus():
    list_bdu_names = get_sheet_list_names()
    dict_bdu_info = get_sheets_info_dict()

    # fmt: off
    df = pd.DataFrame(
        columns=["Fecha","Matrícula","Peso","Origen","Destino","Remito","Pago a",]
    )
    # fmt: on

    for bdu_name in list_bdu_names:
        logger.info("processing the form {bdu_name}".format(bdu_name=bdu_name))
        bdu = get_input_path_file_xslx() + bdu_name
        sheet = dict_bdu_info[bdu_name]["sheet_name"]
        use_cols = dict_bdu_info[bdu_name]["usecols"]  # Example "B,D,G,M,O,P,W,Y"
        bdu_read = read_data_frame(bdu, sheet, use_cols)
        # fmt: off
        bdu_read.columns = ["Fecha","Matrícula","Peso","Origen","Destino","Remito","Pago a",]
        # fmt: on
        bdu_read["Matrícula"] = bdu_read["Matrícula"].str.upper()
        df = pd.concat([df, bdu_read])  # , ignore_index=True)
        logger.info(
            "{value} rows have been added to the travel dataframe".format(
                value=bdu_read.shape[0]
            )
        )
    df["Remito"] = df["Remito"].apply(str)
    return df
