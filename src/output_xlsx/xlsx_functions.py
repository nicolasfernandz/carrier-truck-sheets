import pandas as pd
from output_xlsx import headers
from utils.utils import get_contractor, is_fuel_present
from carrier.rate import get_tree_field_name
from config import CONFIG

contractor = get_contractor()


def write_header(writer, sheetName, index_row):
    if contractor == "EDI":
        data = {CONFIG["header_A"]["key"]: CONFIG["header_A"]["value"]}
        df_header = pd.DataFrame(data, columns=[CONFIG["header_A"]["key"]])
        df_header.to_excel(
            writer, sheet_name=sheetName, index=False, startrow=index_row, startcol=0
        )
        index_row += df_header[CONFIG["header_A"]["key"]].count() + 1
    else:
        data = {CONFIG["header_B"]["key"]: CONFIG["header_B"]["value"]}
        df_header = pd.DataFrame(data, columns=[CONFIG["header_B"]["key"]])
        df_header.to_excel(
            writer, sheet_name=sheetName, index=False, startrow=index_row, startcol=0
        )
        index_row += df_header[CONFIG["header_B"]["key"]].count() + 1

    workbook = writer.book
    worksheet = writer.sheets[sheetName]
    worksheet.set_zoom(85)

    # Write the column header with the defined format.
    for col_num, value in enumerate(df_header.columns.values):
        worksheet.write(0, col_num, value, headers.get_header_cell_format(workbook))
        worksheet.write(
            0, 7, value, headers.get_header_cell_format_fo_last_col(workbook)
        )

    # Write the row values with the defined format.
    for row in df_header.iterrows():
        worksheet.write(
            1, 0, row[1][0], headers.get_header_cell_format_for_merged_cols(workbook)
        )
        worksheet.write(
            1, 1, "", headers.get_header_cell_format_for_merged_cols(workbook)
        )
        worksheet.write(
            1, 2, "", headers.get_header_cell_format_for_merged_cols(workbook)
        )
        worksheet.write(
            1, 3, "", headers.get_header_cell_format_for_merged_cols(workbook)
        )
        worksheet.write(
            1, 4, "", headers.get_header_cell_format_for_merged_cols(workbook)
        )
        worksheet.write(
            1, 5, "", headers.get_header_cell_format_for_merged_cols(workbook)
        )
        worksheet.write(
            1, 6, "", headers.get_header_cell_format_for_merged_cols(workbook)
        )
        worksheet.write(
            1, 7, "", headers.get_header_cell_format_for_merged_cols(workbook)
        )

    worksheet.merge_range("A1:H1", "")
    worksheet.merge_range("A2:H2", "")

    return index_row


def write_carrier(writer, sheetName, index_row, df_carrier_rut):
    data = {
        "Etiqueta": ["Fletero", "Rut"],
        "Valor": [df_carrier_rut.iloc[0, 0], df_carrier_rut.iloc[0, 1]],
    }
    df_rut = pd.DataFrame(data, columns=["Etiqueta", "Valor"])
    df_rut.to_excel(
        writer,
        sheet_name=sheetName,
        header=False,
        index=False,
        startrow=index_row,
        startcol=0,
    )

    workbook = writer.book
    worksheet = writer.sheets[sheetName]
    format_bold_carrier = workbook.add_format({"bold": True})

    # write the row values with the defined format
    for iter, row in enumerate(df_rut.iterrows()):
        worksheet.write(index_row + iter, 0, row[1][0])
        worksheet.write(index_row + iter, 1, row[1][1], format_bold_carrier)

    index_row += df_rut["Etiqueta"].count() + 1
    return index_row, workbook, worksheet


def write_total_exempt(
    writer, sheetName, index_row, workbook, worksheet, total_excempt
):
    data = {
        "Etiqueta": [
            "TOTAL FACTURA EXENTA - (Asimilado a exportacion ART. 34 DEC 220/998)"
        ],
        "B": ["B"],
        "C": ["C"],
        "D": ["D"],
        "Valor": [total_excempt],
    }
    df_total_exempt = pd.DataFrame(data, columns=["Etiqueta", "B", "C", "D", "Valor"])
    df_total_exempt.to_excel(
        writer,
        sheet_name=sheetName,
        header=False,
        index=False,
        startrow=index_row,
        startcol=0,
    )
    index_row += df_total_exempt["Etiqueta"].count() + 1

    for iter, row in enumerate(df_total_exempt.iterrows()):
        worksheet.write(
            index_row - 2,
            0,
            row[1][0],
            headers.get_cell_format(
                workbook,
                color="",
                top=2,
                left=2,
                right=0,
                bottom=2,
                center=True,
                align_left=True,
                liquid_de=contractor,
            ),
        )
        worksheet.set_row(index_row - 2, 27)
        worksheet.write(
            index_row - 2,
            1,
            row[1][1],
            headers.get_cell_format(
                workbook, color="", top=2, left=2, right=0, bottom=2
            ),
        )
        worksheet.write(
            index_row - 2,
            2,
            row[1][2],
            headers.get_cell_format(
                workbook, color="", top=2, left=2, right=0, bottom=2
            ),
        )
        worksheet.write(
            index_row - 2,
            3,
            row[1][3],
            headers.get_cell_format(
                workbook, color="", top=2, left=2, right=0, bottom=2
            ),
        )
        worksheet.write(
            index_row - 2,
            4,
            row[1][4],
            headers.get_cell_format(
                workbook,
                color="",
                top=2,
                left=0,
                right=2,
                bottom=2,
                center=True,
                align_right=True,
                liquid_de=contractor,
            ),
        )

    worksheet = writer.sheets[sheetName]
    _range = "A" + str(index_row - 1) + ":D" + str(index_row - 1)
    worksheet.merge_range(_range, "")
    return index_row


def write_sub_total_tax(writer, sheetName, index_row, workbook, worksheet, total_tax):
    data = {
        "Etiqueta": ["SUBTOTAL FACTURA CON IVA"],
        "B": ["B"],
        "C": ["C"],
        "D": ["D"],
        "Valor": [total_tax],
    }
    df_total_tax = pd.DataFrame(data, columns=["Etiqueta", "B", "C", "D", "Valor"])
    df_total_tax.to_excel(
        writer,
        sheet_name=sheetName,
        header=False,
        index=False,
        startrow=index_row,
        startcol=0,
    )
    index_row += 1

    for iter, row in enumerate(df_total_tax.iterrows()):
        worksheet.write(
            index_row - 1,
            0,
            row[1][0],
            headers.get_cell_format(
                workbook,
                color="",
                top=2,
                left=2,
                right=0,
                bottom=1,
                liquid_de=contractor,
            ),
        )
        worksheet.write(
            index_row - 1,
            1,
            row[1][1],
            headers.get_cell_format(
                workbook, color="", top=2, left=2, right=0, bottom=1
            ),
        )
        worksheet.write(
            index_row - 1,
            2,
            row[1][2],
            headers.get_cell_format(
                workbook, color="", top=2, left=2, right=0, bottom=1
            ),
        )
        worksheet.write(
            index_row - 1,
            3,
            row[1][3],
            headers.get_cell_format(
                workbook, color="", top=2, left=2, right=0, bottom=1
            ),
        )
        worksheet.write(
            index_row - 1,
            4,
            row[1][4],
            headers.get_cell_format(
                workbook,
                color="",
                top=2,
                left=0,
                right=2,
                bottom=1,
                liquid_de=contractor,
            ),
        )

    worksheet = writer.sheets[sheetName]

    _range = "A" + str(index_row) + ":D" + str(index_row)
    worksheet.merge_range(_range, "")
    return index_row


def write_tax(writer, sheetName, index_row, workbook, worksheet, tax):
    data = {"Etiqueta": ["IVA"], "B": ["B"], "C": ["C"], "D": ["D"], "Valor": [tax]}
    df_tax = pd.DataFrame(data, columns=["Etiqueta", "B", "C", "D", "Valor"])
    df_tax.to_excel(
        writer,
        sheet_name=sheetName,
        header=False,
        index=False,
        startrow=index_row,
        startcol=0,
    )
    index_row += 1

    for iter, row in enumerate(df_tax.iterrows()):
        worksheet.write(
            index_row - 1,
            0,
            row[1][0],
            headers.get_cell_format(
                workbook,
                color="",
                top=1,
                left=2,
                right=0,
                bottom=1,
                liquid_de=contractor,
            ),
        )
        worksheet.write(
            index_row - 1,
            1,
            row[1][1],
            headers.get_cell_format(
                workbook, color="", top=1, left=2, right=0, bottom=1
            ),
        )
        worksheet.write(
            index_row - 1,
            2,
            row[1][2],
            headers.get_cell_format(
                workbook, color="", top=1, left=2, right=0, bottom=1
            ),
        )
        worksheet.write(
            index_row - 1,
            3,
            row[1][3],
            headers.get_cell_format(
                workbook, color="", top=1, left=2, right=0, bottom=1
            ),
        )
        worksheet.write(
            index_row - 1,
            4,
            row[1][4],
            headers.get_cell_format(
                workbook,
                color="",
                top=1,
                left=0,
                right=2,
                bottom=1,
                liquid_de=contractor,
            ),
        )

    worksheet = writer.sheets[sheetName]

    _range = "A" + str(index_row) + ":D" + str(index_row)
    worksheet.merge_range(_range, "")
    return index_row


def write_total_tax(writer, sheetName, index_row, workbook, worksheet, total_bill_tax):
    data = {
        "Etiqueta": ["TOTAL FACTURA CON IVA"],
        "B": ["B"],
        "C": ["C"],
        "D": ["D"],
        "Valor": [total_bill_tax],
    }
    df_total_bill_tax = pd.DataFrame(data, columns=["Etiqueta", "B", "C", "D", "Valor"])
    df_total_bill_tax.to_excel(
        writer,
        sheet_name=sheetName,
        header=False,
        index=False,
        startrow=index_row,
        startcol=0,
    )
    index_row += df_total_bill_tax["Etiqueta"].count() + 2

    for iter, row in enumerate(df_total_bill_tax.iterrows()):
        worksheet.write(
            index_row - 3,
            0,
            row[1][0],
            headers.get_cell_format(
                workbook,
                color="",
                top=1,
                left=2,
                right=0,
                bottom=2,
                liquid_de=contractor,
            ),
        )
        worksheet.write(
            index_row - 3,
            1,
            row[1][1],
            headers.get_cell_format(
                workbook, color="", top=1, left=2, right=0, bottom=2
            ),
        )
        worksheet.write(
            index_row - 3,
            2,
            row[1][2],
            headers.get_cell_format(
                workbook, color="", top=1, left=2, right=0, bottom=2
            ),
        )
        worksheet.write(
            index_row - 3,
            3,
            row[1][3],
            headers.get_cell_format(
                workbook, color="", top=1, left=2, right=0, bottom=2
            ),
        )
        worksheet.write(
            index_row - 3,
            4,
            row[1][4],
            headers.get_cell_format(
                workbook,
                color="",
                top=1,
                left=0,
                right=2,
                bottom=2,
                liquid_de=contractor,
            ),
        )

    worksheet = writer.sheets[sheetName]

    _range = "A" + str(index_row - 2) + ":D" + str(index_row - 2)
    worksheet.merge_range(_range, "")
    return index_row


def write_retention_tax(
    writer, sheetName, index_row, workbook, worksheet, retention_tax
):
    data = {
        "Etiqueta": ["RETENCION DE IVA"],
        "B": ["B"],
        "C": ["C"],
        "D": ["D"],
        "Valor": [retention_tax],
    }
    df_retention_tax = pd.DataFrame(data, columns=["Etiqueta", "B", "C", "D", "Valor"])
    df_retention_tax.to_excel(
        writer,
        sheet_name=sheetName,
        header=False,
        index=False,
        startrow=index_row,
        startcol=0,
    )
    index_row += 1

    for iter, row in enumerate(df_retention_tax.iterrows()):
        worksheet.write(
            index_row - 1,
            0,
            row[1][0],
            headers.get_cell_format(workbook, "grey", top=2, left=2, right=0, bottom=1),
        )
        worksheet.write(
            index_row - 1,
            1,
            row[1][1],
            headers.get_cell_format(workbook, "grey", top=2, left=2, right=0, bottom=1),
        )
        worksheet.write(
            index_row - 1,
            2,
            row[1][2],
            headers.get_cell_format(workbook, "grey", top=2, left=2, right=0, bottom=1),
        )
        worksheet.write(
            index_row - 1,
            3,
            row[1][3],
            headers.get_cell_format(workbook, "grey", top=2, left=2, right=0, bottom=1),
        )
        worksheet.write(
            index_row - 1,
            4,
            row[1][4],
            headers.get_cell_format(workbook, "grey", top=2, left=0, right=2, bottom=1),
        )

    worksheet = writer.sheets[sheetName]

    _range = "A" + str(index_row) + ":D" + str(index_row)
    worksheet.merge_range(_range, "")
    return index_row


def write_credit_certificate(
    writer, sheetName, index_row, workbook, worksheet, credit_certificate
):
    data = {
        "Etiqueta": ["CERTIFICADO DE CREDITO"],
        "B": ["B"],
        "C": ["C"],
        "D": ["D"],
        "Valor": [credit_certificate],
    }
    df_credit_certificate = pd.DataFrame(
        data, columns=["Etiqueta", "B", "C", "D", "Valor"]
    )
    df_credit_certificate.to_excel(
        writer,
        sheet_name=sheetName,
        header=False,
        index=False,
        startrow=index_row,
        startcol=0,
    )
    index_row += df_credit_certificate["Etiqueta"].count() + 1

    for iter, row in enumerate(df_credit_certificate.iterrows()):
        worksheet.write(
            index_row - 2,
            0,
            row[1][0],
            headers.get_cell_format(workbook, "grey", top=1, left=2, right=0, bottom=2),
        )
        worksheet.write(
            index_row - 2,
            1,
            row[1][1],
            headers.get_cell_format(workbook, "grey", top=1, left=2, right=0, bottom=2),
        )
        worksheet.write(
            index_row - 2,
            2,
            row[1][2],
            headers.get_cell_format(workbook, "grey", top=1, left=2, right=0, bottom=2),
        )
        worksheet.write(
            index_row - 2,
            3,
            row[1][3],
            headers.get_cell_format(workbook, "grey", top=1, left=2, right=0, bottom=2),
        )
        worksheet.write(
            index_row - 2,
            4,
            row[1][4],
            headers.get_cell_format(workbook, "grey", top=1, left=0, right=2, bottom=2),
        )

    worksheet = writer.sheets[sheetName]

    _range = "A" + str(index_row - 1) + ":D" + str(index_row - 1)
    worksheet.merge_range(_range, "")
    return index_row


def write_total_fuel_consumption(
    writer, sheetName, index_row, workbook, worksheet, total_to_be_deducted
):
    data = {
        "Etiqueta": ["DESCUENTO COMBUSTIBLE"],
        "B": ["B"],
        "C": ["C"],
        "D": ["D"],
        "Valor": [total_to_be_deducted],
    }
    df_fuel_discount = pd.DataFrame(data, columns=["Etiqueta", "B", "C", "D", "Valor"])
    df_fuel_discount.to_excel(
        writer,
        sheet_name=sheetName,
        header=False,
        index=False,
        startrow=index_row,
        startcol=0,
    )
    index_row += 1

    for iter, row in enumerate(df_fuel_discount.iterrows()):
        worksheet.write(
            index_row - 1,
            0,
            row[1][0],
            headers.get_cell_format(
                workbook,
                color="orange",
                top=2,
                left=2,
                right=0,
                bottom=2,
                size=11,
                liquid_de=contractor,
            ),
        )
        worksheet.write(
            index_row - 1,
            1,
            row[1][1],
            headers.get_cell_format(
                workbook, color="orange", top=2, left=2, right=0, bottom=2, size=11
            ),
        )
        worksheet.write(
            index_row - 1,
            2,
            row[1][2],
            headers.get_cell_format(
                workbook, color="orange", top=2, left=2, right=0, bottom=2, size=11
            ),
        )
        worksheet.write(
            index_row - 1,
            3,
            row[1][3],
            headers.get_cell_format(
                workbook, color="orange", top=2, left=2, right=0, bottom=2, size=11
            ),
        )
        worksheet.write(
            index_row - 1,
            4,
            row[1][4],
            headers.get_cell_format(
                workbook,
                color="orange",
                top=2,
                left=0,
                right=2,
                bottom=2,
                size=11,
                numformat="$ #,##0",
                liquid_de=contractor,
            ),
        )  # , numformat='- #,##0'))

    _range = "A" + str(index_row) + ":D" + str(index_row)
    worksheet.merge_range(_range, "")
    return index_row


def write_total_to_be_charged(
    writer, sheetName, index_row, workbook, worksheet, total_to_be_charged
):
    data = {
        "Etiqueta": ["TOTAL A COBRAR"],
        "B": ["B"],
        "C": ["C"],
        "D": ["D"],
        "Valor": [total_to_be_charged],
    }
    df_total_to_be_charged = pd.DataFrame(
        data, columns=["Etiqueta", "B", "C", "D", "Valor"]
    )
    df_total_to_be_charged.to_excel(
        writer,
        sheet_name=sheetName,
        header=False,
        index=False,
        startrow=index_row,
        startcol=0,
    )
    index_row += 1

    for iter, row in enumerate(df_total_to_be_charged.iterrows()):
        worksheet.write(
            index_row - 1,
            0,
            row[1][0],
            headers.get_cell_format(
                workbook,
                color="",
                top=2,
                left=2,
                right=0,
                bottom=2,
                size=14,
                liquid_de=contractor,
            ),
        )
        worksheet.write(
            index_row - 1,
            1,
            row[1][1],
            headers.get_cell_format(
                workbook, color="", top=2, left=2, right=0, bottom=2, size=14
            ),
        )
        worksheet.write(
            index_row - 1,
            2,
            row[1][2],
            headers.get_cell_format(
                workbook, color="", top=2, left=2, right=0, bottom=2, size=14
            ),
        )
        worksheet.write(
            index_row - 1,
            3,
            row[1][3],
            headers.get_cell_format(
                workbook, color="", top=2, left=2, right=0, bottom=2, size=14
            ),
        )
        worksheet.write(
            index_row - 1,
            4,
            row[1][4],
            headers.get_cell_format(
                workbook,
                color="",
                top=2,
                left=0,
                right=2,
                bottom=2,
                size=14,
                liquid_de=contractor,
            ),
        )

    _range = "A" + str(index_row) + ":D" + str(index_row)
    worksheet.merge_range(_range, "")
    return index_row


def write_exempt_forests(vec_forest, df_carrier, writer, sheetName, index_row):
    total_exempt = 0.0
    for i in range(0, len(vec_forest)):
        forest = vec_forest[i]
        # get the list of trips of the forest we are processing from the carrier
        df_carrier_forest = df_carrier[df_carrier["Origen"] == forest]
        # the forest with TAX, when applying the filter should not bring carriers
        if df_carrier_forest["Origen"].count() == 0:
            continue

        # get the list of destinations forests for that origin forest trips
        vec_destination_forest = df_carrier_forest["Destino"].unique()

        for i in range(0, len(vec_destination_forest)):
            destination_forest = vec_destination_forest[i]

            # filter with from origin and destination
            df_carrier_forest = df_carrier[
                (df_carrier["Origen"] == forest)
                & (df_carrier["Destino"] == destination_forest)
            ]

            destination = df_carrier_forest.Destino.iloc[0]
            # add the totals
            df_total_forest = df_carrier_forest.sum(numeric_only=True)
            df_total_forest["$/tonelada"] = df_carrier_forest["$/tonelada"].mean()

            df_total_forest = df_total_forest.iloc[[0, 1, 2]]

            df_carrier_forest = df_carrier_forest.sort_values(
                by=["Fecha llegada en balanza"]
            )
            df_carrier_forest["Fecha llegada en balanza"] = pd.to_datetime(
                df_carrier_forest["Fecha llegada en balanza"]
            ).dt.strftime("%d/%m/%Y")

            total_exempt += df_total_forest["$ Totales"]

            # concat all the trips of the forest with their totals
            # bdu_total = df_transportista_monte.append(df_monte_totales, ignore_index=True)
            bdu_total = pd.concat(
                [df_carrier_forest, df_total_forest.to_frame().T], ignore_index=True
            )

            if bdu_total["Fecha llegada en balanza"].count() > 0:
                bdu_total.to_excel(
                    writer,
                    sheet_name=sheetName,
                    index=False,
                    startrow=index_row,
                    startcol=0,
                )

            workbook = writer.book
            worksheet = writer.sheets[sheetName]
            worksheet.set_column("A:A", 18)
            worksheet.set_column("B:B", 13)
            worksheet.set_column("C:C", 13)
            worksheet.set_column("D:D", 13)
            worksheet.set_column("E:E", 13)
            worksheet.set_column("F:F", 22)
            worksheet.set_column("G:G", 18)
            worksheet.set_column("J:J", 13)
            money_fmt = workbook.add_format(
                {
                    "valign": "top",
                    "text_wrap": True,
                }
            )
            worksheet.set_column("H:I", 15, money_fmt)

            # write the column headers with the defined format
            for col_num, value in enumerate(bdu_total.columns.values):
                if col_num == 2 or col_num == 3 or col_num == 4:
                    worksheet.write(
                        index_row,
                        col_num,
                        value,
                        headers.get_cell_format(workbook, color="grey", center=True),
                    )
                else:
                    worksheet.write(
                        index_row,
                        col_num,
                        value,
                        headers.get_cell_format(workbook, color="grey", center=True),
                    )

            index_row += bdu_total["$ Totales"].count() + 2

            worksheet.write(
                index_row - 2,
                0,
                get_tree_field_name(forest, destination),
                headers.get_cell_format(
                    workbook, top=0, left=0, right=0, bottom=2, liquid_de=contractor
                ),
            )
            worksheet.write(
                index_row - 2,
                1,
                "",
                headers.get_cell_format(
                    workbook, top=0, left=0, right=0, bottom=2, liquid_de=contractor
                ),
            )

            _range = "A" + str(index_row - 1) + ":B" + str(index_row - 1)
            worksheet.merge_range(_range, "")

            for iter, row in enumerate(df_total_forest.to_frame().iterrows()):
                if row[0] == "Peso neto":
                    worksheet.write(
                        index_row - 2,
                        2 + iter,
                        row[1][0],
                        headers.get_cell_format(
                            workbook,
                            top=0,
                            left=0,
                            right=0,
                            bottom=2,
                            setnumformat=False,
                            liquid_de=contractor,
                        ),
                    )
                else:
                    worksheet.write(
                        index_row - 2,
                        2 + iter,
                        row[1][0],
                        headers.get_cell_format(
                            workbook,
                            top=0,
                            left=0,
                            right=0,
                            bottom=2,
                            liquid_de=contractor,
                        ),
                    )

    return index_row, total_exempt


def write_forests_with_tax(vec_forest, df_carrier, writer, sheetName, index_row):
    total_tax = 0.0
    # iterate the carrier's vector of forests
    for i in range(0, len(vec_forest)):
        forest = vec_forest[i]
        # get the list of trips of the forest we are processing from the carrier
        df_carrier_forest = df_carrier[df_carrier["Origen"] == forest]

        # the forest with tax, when applying the filter should not bring Carriers
        if df_carrier_forest["Origen"].count() == 0:
            continue

        # get the list of destinations forests for that origin forest trips
        vec_destination_forest = df_carrier_forest["Destino"].unique()

        for i in range(0, len(vec_destination_forest)):
            destination_forest = vec_destination_forest[i]

            # filter with from origin and destination
            df_carrier_forest = df_carrier[
                (df_carrier["Origen"] == forest)
                & (df_carrier["Destino"] == destination_forest)
            ]

            destination = df_carrier_forest.Destino.iloc[0]

            # add the totals
            df_total_forest = df_carrier_forest.sum(numeric_only=True)
            df_total_forest["$/tonelada"] = df_carrier_forest["$/tonelada"].mean()

            df_total_forest = df_total_forest.iloc[[0, 1, 2]]

            df_carrier_forest = df_carrier_forest.sort_values(
                by=["Fecha llegada en balanza"]
            )
            df_carrier_forest["Fecha llegada en balanza"] = pd.to_datetime(
                df_carrier_forest["Fecha llegada en balanza"]
            ).dt.strftime("%d/%m/%Y")

            total_tax += df_total_forest["$ Totales"]

            # concat all the trips of the forest with their totals
            # bdu_total = df_transportista_monte.append(df_monte_totales, ignore_index=True)
            bdu_total = pd.concat(
                [df_carrier_forest, df_total_forest.to_frame().T], ignore_index=True
            )

            bdu_total = bdu_total[
                [
                    "Fecha llegada en balanza",
                    "Camión",
                    "Peso neto",
                    "$/tonelada",
                    "$ Totales",
                    "Origen",
                    "Destino",
                    "Remito",
                ]
            ]
            bdu_total.columns = [
                "Fecha llegada en balanza",
                "Camión",
                "Peso neto",
                "$/tonelada",
                "$ Totales",
                "Origen",
                "Destino",
                "Remito",
            ]

            if bdu_total["Fecha llegada en balanza"].count() > 0:
                bdu_total.to_excel(
                    writer,
                    sheet_name=sheetName,
                    index=False,
                    startrow=index_row,
                    startcol=0,
                )

            workbook = writer.book
            worksheet = writer.sheets[sheetName]
            worksheet.set_column("A:A", 18)
            worksheet.set_column("B:B", 13)
            worksheet.set_column("C:C", 13)
            worksheet.set_column("D:D", 13)
            worksheet.set_column("E:E", 13)
            worksheet.set_column("F:F", 22)
            worksheet.set_column("G:G", 18)
            worksheet.set_column("J:J", 13)
            money_fmt = workbook.add_format(
                {
                    "valign": "top",
                    "text_wrap": True,
                }
            )
            worksheet.set_column("H:I", 15, money_fmt)

            # write the column headers with the defined format.
            for col_num, value in enumerate(bdu_total.columns.values):
                if col_num == 2 or col_num == 3 or col_num == 4:
                    worksheet.write(
                        index_row,
                        col_num,
                        value,
                        headers.get_cell_format(workbook, color="grey", center=True),
                    )
                else:
                    worksheet.write(
                        index_row,
                        col_num,
                        value,
                        headers.get_cell_format(workbook, color="grey", center=True),
                    )

            index_row += bdu_total["$ Totales"].count() + 2

            worksheet.write(
                index_row - 2,
                0,
                get_tree_field_name(forest, destination),
                headers.get_cell_format(
                    workbook, top=0, left=0, right=0, bottom=2, liquid_de=contractor
                ),
            )
            worksheet.write(
                index_row - 2,
                1,
                "",
                headers.get_cell_format(
                    workbook, top=0, left=0, right=0, bottom=2, liquid_de=contractor
                ),
            )

            _range = "A" + str(index_row - 1) + ":B" + str(index_row - 1)
            worksheet.merge_range(_range, "")

            for iter, row in enumerate(df_total_forest.to_frame().iterrows()):
                if row[0] == "Peso neto":
                    worksheet.write(
                        index_row - 2,
                        2 + iter,
                        row[1][0],
                        headers.get_cell_format(
                            workbook,
                            top=0,
                            left=0,
                            right=0,
                            bottom=2,
                            setnumformat=False,
                            liquid_de=contractor,
                        ),
                    )
                else:
                    worksheet.write(
                        index_row - 2,
                        2 + iter,
                        row[1][0],
                        headers.get_cell_format(
                            workbook,
                            top=0,
                            left=0,
                            right=0,
                            bottom=2,
                            liquid_de=contractor,
                        ),
                    )

    return index_row, total_tax


def write_trips(
    df_truck_id_with_rates,
    df_transporters_rut,
    df_fuel,
    carrier,
    writer,
    df_payment_summary,
):
    # get all trips from the Carrier
    df_carrier = df_truck_id_with_rates[
        df_truck_id_with_rates["Transportista"] == carrier
    ]

    if not df_carrier.empty:
        # get the list of forests from the carrier we are processing
        vec_forest = df_carrier["Origen"].unique()
        if len(vec_forest) > 0:
            index_row = 0
            df_carrier_rut = df_transporters_rut[
                df_transporters_rut["Transportista"] == carrier
            ]

            rut = df_carrier_rut.iloc[0, 1]
            rut = rut.replace(" ", "")
            banco = df_carrier_rut.iloc[0, 2]
            tipo_cuenta = df_carrier_rut.iloc[0, 3]
            cuenta = df_carrier_rut.iloc[0, 4]
            anterior = df_carrier_rut.iloc[0, 5]

            if is_fuel_present():
                df_fuel_loads = df_fuel[df_fuel["Transportista"] == carrier]

            sheetName = carrier[:30]

            index_row = write_header(writer, sheetName, index_row)
            index_row, workbook, worksheet = write_carrier(
                writer, sheetName, index_row, df_carrier_rut
            )

            df_carrier = df_carrier[
                [
                    "Fecha",
                    "Matrícula",
                    "Peso",
                    "Tarifa_Sub",
                    "Total",
                    "Origen",
                    "Destino",
                    "Remito",
                    "IVA",
                ]
            ]
            df_carrier.columns = [
                "Fecha llegada en balanza",
                "Camión",
                "Peso neto",
                "$/tonelada",
                "$ Totales",
                "Origen",
                "Destino",
                "Remito",
                "IVA",
            ]

            df_carrier_excempt = df_carrier[(df_carrier["IVA"] == 0)]

            index_row, total_excempt = write_exempt_forests(
                vec_forest, df_carrier_excempt, writer, sheetName, index_row
            )

            # write total exempt if there is travels
            if total_excempt > 0.0:
                index_row = write_total_exempt(
                    writer, sheetName, index_row, workbook, worksheet, total_excempt
                )

            df_carrier_tax = df_carrier[(df_carrier["IVA"] == 22)]

            index_row, total_tax = write_forests_with_tax(
                vec_forest, df_carrier_tax, writer, sheetName, index_row
            )

            if total_tax > 0.0:
                index_row = write_sub_total_tax(
                    writer, sheetName, index_row, workbook, worksheet, total_tax
                )
            else:
                index_row = write_sub_total_tax(
                    writer, sheetName, index_row, workbook, worksheet, "-"
                )

            tax_iva = total_tax * 0.22
            total_bill_with_tax = total_tax + tax_iva
            retention_tax_iva = tax_iva * 0.6
            credit_certificate = tax_iva - retention_tax_iva
            total_to_be_charged = (
                total_excempt
                + total_bill_with_tax
                - retention_tax_iva
                - credit_certificate
            )

            if total_tax > 0.0:
                index_row = write_tax(
                    writer, sheetName, index_row, workbook, worksheet, tax_iva
                )
            else:
                index_row = write_tax(
                    writer, sheetName, index_row, workbook, worksheet, "-"
                )

            if total_bill_with_tax > 0.0:
                index_row = write_total_tax(
                    writer,
                    sheetName,
                    index_row,
                    workbook,
                    worksheet,
                    total_bill_with_tax,
                )
            else:
                index_row = write_total_tax(
                    writer, sheetName, index_row, workbook, worksheet, "-"
                )

            if retention_tax_iva > 0.0:
                index_row = write_retention_tax(
                    writer,
                    sheetName,
                    index_row,
                    workbook,
                    worksheet,
                    retention_tax_iva * -1,
                )
            else:
                index_row = write_retention_tax(
                    writer, sheetName, index_row, workbook, worksheet, "-"
                )

            if credit_certificate > 0.0:
                index_row = write_credit_certificate(
                    writer,
                    sheetName,
                    index_row,
                    workbook,
                    worksheet,
                    credit_certificate * -1,
                )
            else:
                index_row = write_credit_certificate(
                    writer, sheetName, index_row, workbook, worksheet, "-"
                )

            index_row_fuel_discount_or_total_to_be_charged = index_row

            index_row += 3

            total_fuel = 0.0

            if is_fuel_present() and df_fuel_loads["Fecha de Emisión"].count() > 0:
                # get the list of service stations where their recharged
                vec_stations = df_fuel_loads["Estacion"].unique()
                for i in range(0, len(vec_stations)):
                    fuel_service_station = vec_stations[i]

                    df_fuel_service_station = df_fuel_loads[
                        df_fuel_loads["Estacion"] == fuel_service_station
                    ]
                    df_fuel_service_station = df_fuel_service_station[
                        ["Fecha de Emisión", "Camión", "Comprobante", "Importe"]
                    ]
                    df_total_load = df_fuel_service_station.sum(numeric_only=True)

                    index_row += 1
                    # df_fuel_with_total = df_combustible_estacion.append(df_cargas_totales, ignore_index=True)
                    df_fuel_with_total = pd.concat(
                        [df_fuel_service_station, df_total_load.to_frame().T],
                        ignore_index=True,
                    )
                    df_fuel_with_total.to_excel(
                        writer,
                        sheet_name=sheetName,
                        index=False,
                        startrow=index_row,
                        startcol=0,
                    )

                    # write the column headers with the defined format
                    for col_num, value in enumerate(df_fuel_with_total.columns.values):
                        worksheet.write(
                            index_row,
                            col_num,
                            value,
                            headers.get_cell_format(
                                workbook, color="grey", center=True
                            ),
                        )

                    index_row += df_fuel_with_total["Fecha de Emisión"].count() + 3

                    worksheet.write(
                        index_row - 2,
                        0,
                        "Total Combustible en " + fuel_service_station,
                        headers.get_cell_format(
                            workbook,
                            top=0,
                            left=0,
                            right=0,
                            bottom=2,
                            liquid_de=contractor,
                        ),
                    )
                    worksheet.write(
                        index_row - 2,
                        1,
                        "",
                        headers.get_cell_format(
                            workbook, top=0, left=0, right=0, bottom=2
                        ),
                    )
                    worksheet.write(
                        index_row - 2,
                        2,
                        "",
                        headers.get_cell_format(
                            workbook, top=0, left=0, right=0, bottom=2
                        ),
                    )
                    worksheet.set_column("C:C", 18)
                    worksheet.write(
                        index_row - 2,
                        3,
                        "",
                        headers.get_cell_format(
                            workbook,
                            top=0,
                            left=0,
                            right=0,
                            bottom=2,
                            liquid_de=contractor,
                        ),
                    )

                    _range = "A" + str(index_row - 1) + ":C" + str(index_row - 1)
                    worksheet.merge_range(_range, "")

                    for iter, row in enumerate(df_total_load.to_frame().iterrows()):
                        total_fuel += row[1][0]
                        worksheet.write(
                            index_row - 2,
                            3,
                            row[1][0],
                            headers.get_cell_format(
                                workbook,
                                top=0,
                                left=0,
                                right=0,
                                bottom=2,
                                liquid_de=contractor,
                            ),
                        )

                if total_fuel > 0.0:
                    index_row_fuel_discount_or_total_to_be_charged = (
                        write_total_fuel_consumption(
                            writer,
                            sheetName,
                            index_row_fuel_discount_or_total_to_be_charged,
                            workbook,
                            worksheet,
                            total_fuel * -1,
                        )
                    )

            if total_fuel > 0.0:
                write_total_to_be_charged(
                    writer,
                    sheetName,
                    index_row_fuel_discount_or_total_to_be_charged + 1,
                    workbook,
                    worksheet,
                    total_to_be_charged - total_fuel,
                )
            else:
                write_total_to_be_charged(
                    writer,
                    sheetName,
                    index_row_fuel_discount_or_total_to_be_charged,
                    workbook,
                    worksheet,
                    total_to_be_charged,
                )

        data = {
            "EMPRESA": [carrier],
            "BANCO": [banco],
            "TIPO CUENTA": [tipo_cuenta],
            "CUENTA": [cuenta],
            "ANTERIOR": [anterior],
            "IMPORTE A GIRAR": [total_to_be_charged - total_fuel],
            "RUT": [rut],
            "RETENCION": [retention_tax_iva],
            "CERT": [credit_certificate],
            "DESC. COMBUSTIBLE": [total_fuel],
        }
        if df_payment_summary is None:
            df_payment_summary = pd.DataFrame(
                data,
                columns=[
                    "EMPRESA",
                    "BANCO",
                    "TIPO CUENTA",
                    "CUENTA",
                    "ANTERIOR",
                    "IMPORTE A GIRAR",
                    "RUT",
                    "RETENCION",
                    "CERT",
                    "DESC. COMBUSTIBLE",
                ],
            )
        else:
            df_new_resumen_pago = pd.DataFrame(
                data,
                columns=[
                    "EMPRESA",
                    "BANCO",
                    "TIPO CUENTA",
                    "CUENTA",
                    "ANTERIOR",
                    "IMPORTE A GIRAR",
                    "RUT",
                    "RETENCION",
                    "CERT",
                    "DESC. COMBUSTIBLE",
                ],
            )
            # df_payment_summary = df_resumen_de_pagos.append(df_new_resumen_pago, ignore_index=True)
            df_payment_summary = pd.concat(
                [df_payment_summary, df_new_resumen_pago], ignore_index=True
            )
        return df_payment_summary
