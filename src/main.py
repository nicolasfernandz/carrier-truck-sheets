import logging
import datetime
import pandas as pd
from utils.utils import is_NaN, get_process_month, get_output_path
from carrier.bdu import get_all_bdus
from carrier.fuel import get_fuel
from carrier.rate import get_bdus_with_rates
from carrier.transporters import get_truck_id_with_rates
from carrier.transporters import get_transporters_rut
from output_xlsx.xlsx_functions import write_trips

# Aplication logger
log_filename = "log_" + ".log"
logger = logging.getLogger()
logging.basicConfig(filename=log_filename, level=logging.DEBUG)


def main():
    """Main script to run the Program"""

    # Get application logger
    logger.info("Start travel processing: {}".format(datetime.datetime.now()))

    month = get_process_month()
    output_path = get_output_path()
    df_bdu = get_all_bdus()
    df_fuel = get_fuel()  # combustible_junio
    df_bdu_with_rates = get_bdus_with_rates(df_bdu)  # df_datos_tarifas
    df_truck_id_with_rates = get_truck_id_with_rates(df_bdu_with_rates)  # joined
    df_transporters_rut = get_transporters_rut()  # df_transportistas_rut
    vec_transporters = df_truck_id_with_rates[
        "Transportista"
    ].unique()  # vec_transportistas

    outputFilePath = output_path + "Liquidaciones_Subcontratados_" + month + ".xlsx"
    writer = pd.ExcelWriter(outputFilePath, engine="xlsxwriter")
    df_payment_summary = None
    for i in range(0, len(vec_transporters)):
        if not is_NaN(vec_transporters[i]):
            df_payment_summary = write_trips(
                df_truck_id_with_rates,
                df_transporters_rut,
                df_fuel,
                vec_transporters[i],
                writer,
                df_payment_summary,
            )
    writer.close()

    outputFilePath = output_path + "Resumen_de_pagos_" + month + ".xlsx"
    writer = pd.ExcelWriter(outputFilePath, engine="xlsxwriter")
    df_payment_summary.to_excel(
        writer, sheet_name="Resumen de pagos", index=False, startrow=0, startcol=0
    )
    writer.close()

    print(df_payment_summary)
    logger.info("The process has been completed")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        logger.error("Failed to execute program: {}".format(str(e), exc_info=True))
        raise
