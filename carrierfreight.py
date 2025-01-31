import logging
import pandas as pd
import os.path

logger = logging.getLogger()


class CarrierFreight:
    """Settlement of carrier trips"""

    def __init__(self, main_graph_str, input_paths):
        """
        :param input_paths: Json Section with inputs paths
        """

        if not isinstance(main_graph_str, dict):
            raise Exception("'main_graph_str' must be a dictionary")

        if not isinstance(input_paths, dict):
            raise Exception("'input_paths' must be a dictionary")

        self.main_graph_str = main_graph_str
        self.input_paths = input_paths

        # xsls paths
        self.input_path_file_xslx = (
            input_paths["input_path_file_xslx"] + input_paths["input_file_directory"]
        )
        # rates
        self.input_path_rate = input_paths["input_path_rates"]
        self.df_rate = None
        # dinama
        self.input_path_dinama = input_paths["dinama"]
        # fuel
        self.input_path_fuel = input_paths["fuel"]
        self.df_fuel = None
        # month
        self.month = input_paths["month"]
        # bdus
        self.bdu = None

    def prepare_data(self):
        """ """
        bdus = self.main_graph_str["bdus"]
        list_bdu_names = bdus["sheets_names"]
        dict_bdu_info = self.main_graph_str["sheets"]

        # trips
        df = pd.DataFrame(
            columns=[
                "Fecha",
                "Matrícula",
                "Peso",
                "Origen",
                "Destino",
                "Remito",
                "Pago a",
            ]
        )
        for bdu_name in list_bdu_names:
            logger.info("processing the form {bdu_name}".format(bdu_name=bdu_name))
            bdu = self.input_path_file_xslx + bdu_name
            sheet = dict_bdu_info[bdu_name]["sheet_name"]
            useCols = dict_bdu_info[bdu_name]["usecols"]  # Example "B,D,G,M,O,P,W,Y"
            bdu_read = pd.read_excel(bdu, sheet_name=sheet, usecols=useCols)
            bdu_read.columns = [
                "Fecha",
                "Matrícula",
                "Peso",
                "Origen",
                "Destino",
                "Remito",
                "Transportista",
            ]
            bdu_read = bdu_read[
                [
                    "Fecha",
                    "Matrícula",
                    "Peso",
                    "Origen",
                    "Destino",
                    "Remito",
                    "Transportista",
                ]
            ]
            bdu_read.columns = [
                "Fecha",
                "Matrícula",
                "Peso",
                "Origen",
                "Destino",
                "Remito",
                "Pago a",
            ]
            bdu_read["Matrícula"] = bdu_read["Matrícula"].str.upper()
            df = pd.concat([df, bdu_read], ignore_index=True)
            logger.info(
                "{value} rows have been added to the travel dataframe".format(
                    value=bdu_read.shape[0]
                )
            )
        df["Remito"] = df["Remito"].apply(str)
        self.bdu = df

        # fuel
        if os.path.isfile(self.input_path_file_xslx + self.input_path_fuel):
            logger.info(
                "processing fuel form {fuel_sheet_name}".format(
                    fuel_sheet_name=self.input_path_fuel
                )
            )
            df_fuel_keys = pd.read_excel(
                self.input_path_file_xslx + self.input_path_fuel, sheet_name=None
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
                df_fuel_read = pd.read_excel(
                    self.input_path_file_xslx + self.input_path_fuel,
                    sheet_name=sheet_name,
                    usecols="A,B,C,D,E,F",
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
            self.df_fuel = df_fuel_concat
            logger.info(
                "total number of rows of the fuel dataframe is {value}".format(
                    value=self.df_fuel.shape[0]
                )
            )
        
        # rates
        rates_read=pd.read_excel(self.input_path_rate, sheet_name="Actual", usecols="A,B,D")
        df_rate = pd.merge(self.bdu, rates_read[['Sector/Pila origen', 'Destino', 'Tarifa_Sub']], how='left',left_on=['Origen','Destino'], right_on=['Sector/Pila origen','Destino'])
        df_rate = df_rate[["Fecha", "Matrícula", "Origen", "Destino", "Remito", "Pago a", "Peso", "Tarifa_Sub"]]
        df_rate['Total'] = df_rate.Peso * df_rate.Tarifa_Sub
        self.df_rate = df_rate
        logger.info(
            "total number of rows of the rate dataframe is {value}".format(
                value=rates_read.shape[0]
            )
        )


    def run(self):
        """
        :return:
        """
        self.prepare_data()
        print(self.bdu)
        print(self.df_fuel)
