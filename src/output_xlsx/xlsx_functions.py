import pandas as pd
from output_xlsx import headers
from utils.utils import get_contractor, is_fuel_present
from carrier.rate import get_tree_field_name
from config import CONFIG

contractor = get_contractor()

def write_header(writer, sheetName, index_row):
    if contractor == "EDI":
        data = {CONFIG['header_A']['key']: CONFIG['header_A']['value']}
        df_titulo = pd.DataFrame (data, columns = [CONFIG['header_A']['key']])
        df_titulo.to_excel(writer, sheet_name=sheetName, index=False, startrow=index_row, startcol=0)
        index_row += df_titulo[CONFIG['header_A']['key']].count() + 1
    else:
        data = {CONFIG['header_B']['key']: CONFIG['header_B']['value']}
        df_titulo = pd.DataFrame (data, columns = [CONFIG['header_B']['key']])
        df_titulo.to_excel(writer, sheet_name=sheetName, index=False, startrow=index_row, startcol=0)
        index_row += df_titulo[CONFIG['header_B']['key']].count() + 1

    workbook = writer.book
    worksheet = writer.sheets[sheetName]
    worksheet.set_zoom(85)
 
    # Write the column header with the defined format.
    for col_num, value in enumerate(df_titulo.columns.values):
        worksheet.write(0, col_num, value, headers.get_header_cell_format(workbook))
        worksheet.write(0, 7, value, headers.get_header_cell_format_fo_last_col(workbook))

    # Write the row values with the defined format.
    for row in df_titulo.iterrows():
        worksheet.write(1, 0, row[1][0], headers.get_header_cell_format_for_merged_cols(workbook))
        worksheet.write(1, 1, '', headers.get_header_cell_format_for_merged_cols(workbook))
        worksheet.write(1, 2, '', headers.get_header_cell_format_for_merged_cols(workbook))
        worksheet.write(1, 3, '', headers.get_header_cell_format_for_merged_cols(workbook))
        worksheet.write(1, 4, '', headers.get_header_cell_format_for_merged_cols(workbook))
        worksheet.write(1, 5, '', headers.get_header_cell_format_for_merged_cols(workbook))
        worksheet.write(1, 6, '', headers.get_header_cell_format_for_merged_cols(workbook))
        worksheet.write(1, 7, '', headers.get_header_cell_format_for_merged_cols(workbook))

    worksheet.merge_range('A1:H1', "")
    worksheet.merge_range('A2:H2', "")

    return index_row

def write_carrier(writer, sheetName, index_row, df_fletero_rut):
    data = {'Etiqueta':['Fletero', 'Rut'], 'Valor':[df_fletero_rut.iloc[0, 0], df_fletero_rut.iloc[0, 1]]}
    df_rut = pd.DataFrame (data, columns = ['Etiqueta','Valor'])
    df_rut.to_excel(writer, sheet_name=sheetName, header=False, index=False, startrow=index_row, startcol=0)

    workbook = writer.book
    worksheet = writer.sheets[sheetName]
    format_bold_fletero = workbook.add_format({ 'bold': True })

    # write the row values with the defined format
    for iter, row in enumerate(df_rut.iterrows()):
        worksheet.write(index_row+iter, 0, row[1][0])
        worksheet.write(index_row+iter, 1, row[1][1], format_bold_fletero)

    index_row += df_rut['Etiqueta'].count() + 1
    return index_row, workbook, worksheet

def write_total_exempt(writer, sheetName, index_row, workbook, worksheet, total_excento):
    data = {'Etiqueta':['TOTAL FACTURA EXENTA - (Asimilado a exportacion ART. 34 DEC 220/998)'],'B':['B'], 'C':['C'], 'D':['D'], 'Valor':[total_excento]}
    df_total_excento = pd.DataFrame (data, columns = ['Etiqueta', 'B', 'C', 'D', 'Valor'])
    df_total_excento.to_excel(writer, sheet_name=sheetName, header=False, index=False, startrow=index_row, startcol=0)
    index_row += df_total_excento['Etiqueta'].count() + 1

    for iter, row in enumerate(df_total_excento.iterrows()):
        worksheet.write(index_row-2, 0, row[1][0], headers.get_cell_format(workbook, color='', top=2, left=2, right=0, bottom=2, center=True, align_left=True, liquid_de=contractor))
        worksheet.set_row(index_row-2, 27)
        worksheet.write(index_row-2, 1, row[1][1], headers.get_cell_format(workbook, color='', top=2, left=2, right=0, bottom=2))
        worksheet.write(index_row-2, 2, row[1][2], headers.get_cell_format(workbook, color='', top=2, left=2, right=0, bottom=2))
        worksheet.write(index_row-2, 3, row[1][3], headers.get_cell_format(workbook, color='', top=2, left=2, right=0, bottom=2))
        worksheet.write(index_row-2, 4, row[1][4], headers.get_cell_format(workbook, color='', top=2, left=0, right=2, bottom=2, center=True, align_right=True, liquid_de=contractor))

    worksheet = writer.sheets[sheetName]
    _range = 'A' + str(index_row-1) + ':D' + str(index_row-1)
    worksheet.merge_range(_range, "")
    return index_row

def write_sub_total_tax(writer, sheetName, index_row, workbook, worksheet, total_iva):
    data = {'Etiqueta':['SUBTOTAL FACTURA CON IVA'], 'B':['B'], 'C':['C'], 'D':['D'], 'Valor':[total_iva]}
    df_total_iva = pd.DataFrame (data, columns = ['Etiqueta', 'B', 'C', 'D', 'Valor'])
    df_total_iva.to_excel(writer, sheet_name=sheetName, header=False, index=False, startrow=index_row, startcol=0)
    index_row += 1

    for iter, row in enumerate(df_total_iva.iterrows()):
        worksheet.write(index_row-1, 0, row[1][0], headers.get_cell_format(workbook, color='', top=2, left=2, right=0, bottom=1, liquid_de=contractor))
        worksheet.write(index_row-1, 1, row[1][1], headers.get_cell_format(workbook, color='', top=2, left=2, right=0, bottom=1))
        worksheet.write(index_row-1, 2, row[1][2], headers.get_cell_format(workbook, color='', top=2, left=2, right=0, bottom=1))
        worksheet.write(index_row-1, 3, row[1][3], headers.get_cell_format(workbook, color='', top=2, left=2, right=0, bottom=1))
        worksheet.write(index_row-1, 4, row[1][4], headers.get_cell_format(workbook, color='', top=2, left=0, right=2, bottom=1, liquid_de=contractor))

    worksheet = writer.sheets[sheetName]

    _range = 'A' + str(index_row) + ':D' + str(index_row)
    worksheet.merge_range(_range, "")
    return index_row

def write_tax(writer, sheetName, index_row, workbook, worksheet, iva):
    data = {'Etiqueta':['IVA'], 'B':['B'], 'C':['C'], 'D':['D'], 'Valor':[iva]}
    df_iva = pd.DataFrame (data, columns = ['Etiqueta', 'B', 'C', 'D', 'Valor'])
    df_iva.to_excel(writer, sheet_name=sheetName, header=False, index=False, startrow=index_row, startcol=0)
    index_row += 1

    for iter, row in enumerate(df_iva.iterrows()):
        worksheet.write(index_row-1, 0, row[1][0], headers.get_cell_format(workbook, color='', top=1, left=2, right=0, bottom=1, liquid_de=contractor))
        worksheet.write(index_row-1, 1, row[1][1], headers.get_cell_format(workbook, color='', top=1, left=2, right=0, bottom=1))
        worksheet.write(index_row-1, 2, row[1][2], headers.get_cell_format(workbook, color='', top=1, left=2, right=0, bottom=1))
        worksheet.write(index_row-1, 3, row[1][3], headers.get_cell_format(workbook, color='', top=1, left=2, right=0, bottom=1))
        worksheet.write(index_row-1, 4, row[1][4], headers.get_cell_format(workbook, color='', top=1, left=0, right=2, bottom=1, liquid_de=contractor))

    worksheet = writer.sheets[sheetName]

    _range = 'A' + str(index_row) + ':D' + str(index_row)
    worksheet.merge_range(_range, "")
    return index_row

def write_total_tax(writer, sheetName, index_row, workbook, worksheet, total_fac_iva):
    data = {'Etiqueta':['TOTAL FACTURA CON IVA'], 'B':['B'], 'C':['C'], 'D':['D'], 'Valor':[total_fac_iva]}
    df_total_fac_iva = pd.DataFrame (data, columns = ['Etiqueta', 'B', 'C', 'D', 'Valor'])
    df_total_fac_iva.to_excel(writer, sheet_name=sheetName, header=False, index=False, startrow=index_row, startcol=0)
    index_row += df_total_fac_iva['Etiqueta'].count() + 2

    for iter, row in enumerate(df_total_fac_iva.iterrows()):
        worksheet.write(index_row-3, 0, row[1][0], headers.get_cell_format(workbook, color='', top=1, left=2, right=0, bottom=2, liquid_de=contractor))
        worksheet.write(index_row-3, 1, row[1][1], headers.get_cell_format(workbook, color='', top=1, left=2, right=0, bottom=2))
        worksheet.write(index_row-3, 2, row[1][2], headers.get_cell_format(workbook, color='', top=1, left=2, right=0, bottom=2))
        worksheet.write(index_row-3, 3, row[1][3], headers.get_cell_format(workbook, color='', top=1, left=2, right=0, bottom=2))
        worksheet.write(index_row-3, 4, row[1][4], headers.get_cell_format(workbook, color='', top=1, left=0, right=2, bottom=2, liquid_de=contractor))

    worksheet = writer.sheets[sheetName]

    _range = 'A' + str(index_row-2) + ':D' + str(index_row-2)
    worksheet.merge_range(_range, "")
    return index_row

def write_retention_tax(writer, sheetName, index_row, workbook, worksheet, retencion_iva):
    data = {'Etiqueta':['RETENCION DE IVA'], 'B':['B'], 'C':['C'], 'D':['D'], 'Valor':[retencion_iva]}
    df_ret_iva = pd.DataFrame (data, columns = ['Etiqueta', 'B', 'C', 'D', 'Valor'])
    df_ret_iva.to_excel(writer, sheet_name=sheetName, header=False, index=False, startrow=index_row, startcol=0)
    index_row += 1

    for iter, row in enumerate(df_ret_iva.iterrows()):
        worksheet.write(index_row-1, 0, row[1][0], headers.get_cell_format(workbook, 'grey', top=2, left=2, right=0, bottom=1))
        worksheet.write(index_row-1, 1, row[1][1], headers.get_cell_format(workbook, 'grey', top=2, left=2, right=0, bottom=1))
        worksheet.write(index_row-1, 2, row[1][2], headers.get_cell_format(workbook, 'grey', top=2, left=2, right=0, bottom=1))
        worksheet.write(index_row-1, 3, row[1][3], headers.get_cell_format(workbook, 'grey', top=2, left=2, right=0, bottom=1))
        worksheet.write(index_row-1, 4, row[1][4], headers.get_cell_format(workbook, 'grey', top=2, left=0, right=2, bottom=1))

    worksheet = writer.sheets[sheetName]

    _range = 'A' + str(index_row) + ':D' + str(index_row)
    worksheet.merge_range(_range, "")
    return index_row

def write_credit_certificate(writer, sheetName, index_row, workbook, worksheet, certificado_de_credito):
    data = {'Etiqueta':['CERTIFICADO DE CREDITO'], 'B':['B'], 'C':['C'], 'D':['D'], 'Valor':[certificado_de_credito]}
    df_cert_credito = pd.DataFrame (data, columns = ['Etiqueta', 'B', 'C', 'D', 'Valor'])
    df_cert_credito.to_excel(writer, sheet_name=sheetName, header=False, index=False, startrow=index_row, startcol=0)
    index_row += df_cert_credito['Etiqueta'].count() + 1

    for iter, row in enumerate(df_cert_credito.iterrows()):
        worksheet.write(index_row-2, 0, row[1][0], headers.get_cell_format(workbook, 'grey', top=1, left=2, right=0, bottom=2))
        worksheet.write(index_row-2, 1, row[1][1], headers.get_cell_format(workbook, 'grey', top=1, left=2, right=0, bottom=2))
        worksheet.write(index_row-2, 2, row[1][2], headers.get_cell_format(workbook, 'grey', top=1, left=2, right=0, bottom=2))
        worksheet.write(index_row-2, 3, row[1][3], headers.get_cell_format(workbook, 'grey', top=1, left=2, right=0, bottom=2))
        worksheet.write(index_row-2, 4, row[1][4], headers.get_cell_format(workbook, 'grey', top=1, left=0, right=2, bottom=2))

    worksheet = writer.sheets[sheetName]

    _range = 'A' + str(index_row-1) + ':D' + str(index_row-1)
    worksheet.merge_range(_range, "")
    return index_row

def write_total_fuel_consumption(writer, sheetName, index_row, workbook, worksheet, total_a_cobrar):
    data = {'Etiqueta':['DESCUENTO COMBUSTIBLE'], 'B':['B'], 'C':['C'], 'D':['D'], 'Valor':[total_a_cobrar]}
    df_desc_combustible = pd.DataFrame (data, columns = ['Etiqueta', 'B', 'C', 'D', 'Valor'])
    df_desc_combustible.to_excel(writer, sheet_name=sheetName, header=False, index=False, startrow=index_row, startcol=0)
    index_row += 1

    for iter, row in enumerate(df_desc_combustible.iterrows()):
        worksheet.write(index_row-1, 0, row[1][0], headers.get_cell_format(workbook, color='orange', top=2, left=2, right=0, bottom=2, size=11, liquid_de=contractor))
        worksheet.write(index_row-1, 1, row[1][1], headers.get_cell_format(workbook, color='orange', top=2, left=2, right=0, bottom=2, size=11))
        worksheet.write(index_row-1, 2, row[1][2], headers.get_cell_format(workbook, color='orange', top=2, left=2, right=0, bottom=2, size=11))
        worksheet.write(index_row-1, 3, row[1][3], headers.get_cell_format(workbook, color='orange', top=2, left=2, right=0, bottom=2, size=11))
        worksheet.write(index_row-1, 4, row[1][4], headers.get_cell_format(workbook, color='orange', top=2, left=0, right=2, bottom=2, size=11, numformat='$ #,##0', liquid_de=contractor)) #, numformat='- #,##0'))

    _range = 'A' + str(index_row) + ':D' + str(index_row)
    worksheet.merge_range(_range, "")
    return index_row

def write_total_to_be_charged(writer, sheetName, index_row, workbook, worksheet, total_a_cobrar):
    data = {'Etiqueta':['TOTAL A COBRAR'], 'B':['B'], 'C':['C'], 'D':['D'], 'Valor':[total_a_cobrar]}
    df_total_a_cobrar = pd.DataFrame (data, columns = ['Etiqueta', 'B', 'C', 'D', 'Valor'])
    df_total_a_cobrar.to_excel(writer, sheet_name=sheetName, header=False, index=False, startrow=index_row, startcol=0)
    index_row += 1

    for iter, row in enumerate(df_total_a_cobrar.iterrows()):
        worksheet.write(index_row-1, 0, row[1][0], headers.get_cell_format(workbook, color='', top=2, left=2, right=0, bottom=2, size=14, liquid_de=contractor))
        worksheet.write(index_row-1, 1, row[1][1], headers.get_cell_format(workbook, color='', top=2, left=2, right=0, bottom=2, size=14))
        worksheet.write(index_row-1, 2, row[1][2], headers.get_cell_format(workbook, color='', top=2, left=2, right=0, bottom=2, size=14))
        worksheet.write(index_row-1, 3, row[1][3], headers.get_cell_format(workbook, color='', top=2, left=2, right=0, bottom=2, size=14))
        worksheet.write(index_row-1, 4, row[1][4], headers.get_cell_format(workbook, color='', top=2, left=0, right=2, bottom=2, size=14, liquid_de=contractor))

    _range = 'A' + str(index_row) + ':D' + str(index_row)
    worksheet.merge_range(_range, "")
    return index_row

def write_exempt_forests(vec_montes, df_transportista, writer, sheetName, index_row):
    total_exempt = 0.0
    for i in range(0, len(vec_montes)):
        monte = vec_montes[i]
        # get the list of trips of the forest we are processing from the carrier
        df_transportista_monte = df_transportista[df_transportista['Origen'] == monte]
        # the forest with TAX, when applying the filter should not bring carriers
        if df_transportista_monte['Origen'].count() == 0:
            continue

        _destino = df_transportista_monte.Destino.iloc[0]
        # add the totals
        df_monte_totales = df_transportista_monte.sum(numeric_only=True)
        df_monte_totales['$/tonelada']= df_transportista_monte["$/tonelada"].mean()

        df_transportista_monte=df_transportista_monte.sort_values(by=['Fecha llegada en balanza'])
        df_transportista_monte["Fecha llegada en balanza"] = pd.to_datetime(df_transportista_monte["Fecha llegada en balanza"]).dt.strftime("%d/%m/%Y")

        total_exempt += df_monte_totales["$ Totales"]

        # concat all the trips of the forest with their totals
        #bdu_total = df_transportista_monte.append(df_monte_totales, ignore_index=True)
        bdu_total = pd.concat([df_transportista_monte, df_monte_totales.to_frame().T], ignore_index=True)

        if(bdu_total['Fecha llegada en balanza'].count() > 0):
            bdu_total.to_excel(writer, sheet_name=sheetName, index=False, startrow=index_row, startcol=0)

        workbook = writer.book
        worksheet = writer.sheets[sheetName]
        worksheet.set_column('A:A', 18)
        worksheet.set_column('B:B', 13)
        worksheet.set_column('C:C', 13)
        worksheet.set_column('D:D', 13)
        worksheet.set_column('E:E', 13)
        worksheet.set_column('F:F', 22)
        worksheet.set_column('G:G', 18)
        worksheet.set_column('J:J', 13)
        money_fmt = workbook.add_format({'valign': 'top', 'text_wrap': True,})
        worksheet.set_column('H:I', 15, money_fmt)

        # write the column headers with the defined format
        for col_num, value in enumerate(bdu_total.columns.values):
            if col_num == 2 or col_num == 3 or col_num == 4 :
                worksheet.write(index_row, col_num, value, headers.get_cell_format(workbook, color='grey', center=True))
            else:
                worksheet.write(index_row, col_num, value, headers.get_cell_format(workbook, color='grey', center=True))

        index_row += bdu_total['$ Totales'].count() + 2

        worksheet.write(index_row-2, 0, get_tree_field_name(monte, _destino), headers.get_cell_format(workbook, top=0, left=0, right=0, bottom=2, liquid_de=contractor))
        worksheet.write(index_row-2, 1, "", headers.get_cell_format(workbook, top=0, left=0, right=0, bottom=2, liquid_de=contractor))

        _range = 'A' + str(index_row-1) + ':B' + str(index_row-1)
        worksheet.merge_range(_range, "")

        for iter, row in enumerate(df_monte_totales.to_frame().iterrows()):
            if row[0] == 'Peso neto':
                worksheet.write(index_row-2, 2+iter, row[1][0], headers.get_cell_format(workbook, top=0, left=0, right=0, bottom=2, setnumformat=False, liquid_de=contractor))
            else:
                worksheet.write(index_row-2, 2+iter, row[1][0], headers.get_cell_format(workbook, top=0, left=0, right=0, bottom=2, liquid_de=contractor))

    return index_row, total_exempt

def write_forests_with_tax(vec_montes, df_transportista, writer, sheetName, index_row):
    total_iva = 0.0
    # iterate the carrier's vector of forests
    for i in range(0, len(vec_montes)):
        monte = vec_montes[i]
        # get the list of trips of the forest we are processing from the carrier
        df_transportista_monte = df_transportista[df_transportista['Origen'] == monte]

        # the forest with tax, when applying the filter should not bring Carriers
        if df_transportista_monte['Origen'].count() == 0:
            continue

        _destino = df_transportista_monte.Destino.iloc[0]

        # add the totals
        df_monte_totales = df_transportista_monte.sum(numeric_only=True)

        df_monte_totales['$/tonelada']= df_transportista_monte["$/tonelada"].mean()
        print(df_monte_totales)

        df_transportista_monte=df_transportista_monte.sort_values(by=['Fecha llegada en balanza'])
        df_transportista_monte["Fecha llegada en balanza"] = pd.to_datetime(df_transportista_monte["Fecha llegada en balanza"]).dt.strftime("%d/%m/%Y")

        total_iva += df_monte_totales["$ Totales"]
    
        # concat all the trips of the forest with their totals
        #bdu_total = df_transportista_monte.append(df_monte_totales, ignore_index=True)
        bdu_total = pd.concat([df_transportista_monte, df_monte_totales.to_frame().T], ignore_index=True)

        print(bdu_total)

        bdu_total = bdu_total[["Fecha llegada en balanza", "Camión", "Peso neto", "$/tonelada", "$ Totales", "Origen", "Destino", "Remito"]]
        bdu_total.columns = ["Fecha llegada en balanza", "Camión", "Peso neto", "$/tonelada", "$ Totales", "Origen", "Destino", "Remito"]

        if(bdu_total['Fecha llegada en balanza'].count() > 0):
            bdu_total.to_excel(writer, sheet_name=sheetName, index=False, startrow=index_row, startcol=0)

        workbook = writer.book
        worksheet = writer.sheets[sheetName]
        worksheet.set_column('A:A', 18)
        worksheet.set_column('B:B', 13)
        worksheet.set_column('C:C', 13)
        worksheet.set_column('D:D', 13)
        worksheet.set_column('E:E', 13)
        worksheet.set_column('F:F', 22)
        worksheet.set_column('G:G', 18)
        worksheet.set_column('J:J', 13)
        money_fmt = workbook.add_format({'valign': 'top', 'text_wrap': True,})
        worksheet.set_column('H:I', 15, money_fmt)

        # write the column headers with the defined format.
        for col_num, value in enumerate(bdu_total.columns.values):
            if col_num == 2 or col_num == 3 or col_num == 4 :
                worksheet.write(index_row, col_num, value, headers.get_cell_format(workbook, color='grey', center=True))
            else:
                worksheet.write(index_row, col_num, value, headers.get_cell_format(workbook, color='grey', center=True))

        index_row += bdu_total['$ Totales'].count() + 2

        worksheet.write(index_row-2, 0, get_tree_field_name(monte, _destino), headers.get_cell_format(workbook, top=0, left=0, right=0, bottom=2, liquid_de=contractor))
        worksheet.write(index_row-2, 1, "", headers.get_cell_format(workbook, top=0, left=0, right=0, bottom=2, liquid_de=contractor))

        _range = 'A' + str(index_row-1) + ':B' + str(index_row-1)
        worksheet.merge_range(_range, "")

        for iter, row in enumerate(df_monte_totales.to_frame().iterrows()):
            if row[0] == 'Peso neto':
                worksheet.write(index_row-2, 2+iter, row[1][0], headers.get_cell_format(workbook, top=0, left=0, right=0, bottom=2, setnumformat=False, liquid_de=contractor))
            else:
                worksheet.write(index_row-2, 2+iter, row[1][0], headers.get_cell_format(workbook, top=0, left=0, right=0, bottom=2, liquid_de=contractor))

    return index_row, total_iva

def write_trips(joined, df_transportistas_rut, combustible_junio, transportista, writer, df_resumen_de_pagos):
    # get all trips from the Carrier
    df_transportista = joined[joined['Transportista'] == transportista]

    if(not df_transportista.empty):
        # get the list of forests from the carrier we are processing
        vec_montes = df_transportista["Origen"].unique()
        if(len(vec_montes) > 0):
            index_row = 0
            df_fletero_rut = df_transportistas_rut[df_transportistas_rut['Transportista'] == transportista]

            rut = df_fletero_rut.iloc[0, 1]
            rut = rut.replace(' ','')
            banco = df_fletero_rut.iloc[0, 2]
            tipo_cuenta = df_fletero_rut.iloc[0, 3]
            cuenta = df_fletero_rut.iloc[0, 4]
            anterior = df_fletero_rut.iloc[0, 5]

            print(transportista + " " + rut + " " + banco + " " + tipo_cuenta + " " + str(cuenta) + " " + str(anterior))

            if is_fuel_present():
                df_cargas_gasoil = combustible_junio[combustible_junio['Transportista'] == transportista]

            sheetName=transportista[:30]

            index_row = write_header(writer, sheetName, index_row)
            index_row, workbook, worksheet = write_carrier(writer, sheetName, index_row, df_fletero_rut)

            df_transportista = df_transportista[["Fecha", "Matrícula", "Peso", "Tarifa_Sub", "Total", "Origen", "Destino", "Remito"]]
            df_transportista.columns = ["Fecha llegada en balanza", "Camión", "Peso neto", "$/tonelada", "$ Totales", "Origen", "Destino", "Remito"]

            df_transportista_excentos = df_transportista[(df_transportista['Destino'] == "MDP") | (df_transportista['Destino'] == "UPM") | (df_transportista['Destino'] == "UPM2")]

            index_row, total_excento = write_exempt_forests(vec_montes, df_transportista_excentos, writer, sheetName, index_row)

            # write total exempt if there is travels
            if total_excento > 0.0:
                index_row = write_total_exempt(writer, sheetName, index_row, workbook, worksheet, total_excento)

            # TO DO - TAX!!!
            df_transportista_iva = df_transportista[(df_transportista['Destino'] == "Kluntex SA") | (df_transportista['Destino'] == "Murchison S.A.") | (df_transportista['Destino'] == "Sierra Verde") | (df_transportista['Destino'] == "Acopio Olimar") | (df_transportista['Destino'] == "Planta Olimar") | (df_transportista['Destino'] == "ACOPIO CERRO CHATO") | (df_transportista['Destino'] == "ACOPIO") | (df_transportista['Destino'] == "Balanza") | (df_transportista['Destino'] == "PLANIR MVDEO") | (df_transportista['Destino'] == "PLANIR PROGRESO") | (df_transportista['Destino'] == "ACOPIO SANTAMARIA") | (df_transportista['Destino'] == "Acopio Santamaria") | (df_transportista['Destino'] == "ACOPIO PLANIR") | (df_transportista['Destino'] == "ACOPIO TACUAREMBÓ") | (df_transportista['Destino'] == "JUANGO") | (df_transportista['Destino'] == "Chipper") | (df_transportista['Destino'] == "DEPOSITO DE PACO") | (df_transportista['Destino'] == "URCEL") | (df_transportista['Destino'] == "FAS") | (df_transportista['Destino'] == "IDALEN S.A.") | (df_transportista['Destino'] == "BALANZA TROZA 5") | (df_transportista['Destino'] == "Transportes José Pedro Varela S. A")]

            print(df_transportista_iva)

            index_row, total_iva = write_forests_with_tax(vec_montes, df_transportista_iva, writer, sheetName, index_row)

            if total_iva > 0.0:
                index_row = write_sub_total_tax(writer, sheetName, index_row, workbook, worksheet, total_iva)
            else:
                index_row = write_sub_total_tax(writer, sheetName, index_row, workbook, worksheet, '-')

            iva = total_iva * 0.22
            total_factura_con_iva = total_iva + iva
            retencion_iva =  iva * 0.6
            certificado_de_credito = iva - retencion_iva
            total_a_cobrar = total_excento + total_factura_con_iva - retencion_iva - certificado_de_credito

            if total_iva > 0.0:
                index_row = write_tax(writer, sheetName, index_row, workbook, worksheet, iva)
            else:
                index_row = write_tax(writer, sheetName, index_row, workbook, worksheet, "-")

            if total_factura_con_iva > 0.0:
                index_row = write_total_tax(writer, sheetName, index_row, workbook, worksheet, total_factura_con_iva)
            else:
                index_row = write_total_tax(writer, sheetName, index_row, workbook, worksheet, "-")

            if retencion_iva > 0.0:
                index_row = write_retention_tax(writer, sheetName, index_row, workbook, worksheet, retencion_iva * -1)
            else:
                index_row = write_retention_tax(writer, sheetName, index_row, workbook, worksheet, "-")

            if certificado_de_credito > 0.0:
                index_row = write_credit_certificate(writer, sheetName, index_row, workbook, worksheet, certificado_de_credito * -1)
            else:
                index_row = write_credit_certificate(writer, sheetName, index_row, workbook, worksheet, "-")

            index_row_descuento_combustible_o_total_a_cobrar = index_row

            index_row += 3

            total_combustible = 0.0

            if(is_fuel_present() and df_cargas_gasoil['Fecha de Emisión'].count() > 0):
                # get the list of service stations where their recharged
                vec_estaciones = df_cargas_gasoil["Estacion"].unique()
                for i in range(0, len(vec_estaciones)):
                    estacion = vec_estaciones[i]

                    df_combustible_estacion = df_cargas_gasoil[df_cargas_gasoil["Estacion"] == estacion]
                    df_combustible_estacion = df_combustible_estacion[['Fecha de Emisión', 'Camión', 'Comprobante', 'Importe']]
                    df_cargas_totales = df_combustible_estacion.sum(numeric_only=True)

                    index_row += 1
                    #df_gasoil_with_total = df_combustible_estacion.append(df_cargas_totales, ignore_index=True)
                    df_gasoil_with_total = pd.concat([df_combustible_estacion, df_cargas_totales.to_frame().T], ignore_index=True)
                    df_gasoil_with_total.to_excel(writer, sheet_name=sheetName, index=False, startrow=index_row, startcol=0)

                    # write the column headers with the defined format
                    for col_num, value in enumerate(df_gasoil_with_total.columns.values):
                                                            worksheet.write(index_row, col_num, value, headers.get_cell_format(workbook, color='grey', center=True))

                    index_row += df_gasoil_with_total['Fecha de Emisión'].count() + 3

                    worksheet.write(index_row-2, 0, 'Total Combustible en ' + estacion, headers.get_cell_format(workbook, top=0, left=0, right=0, bottom=2, liquid_de=contractor))
                    worksheet.write(index_row-2, 1, "", headers.get_cell_format(workbook, top=0, left=0, right=0, bottom=2))
                    worksheet.write(index_row-2, 2, "", headers.get_cell_format(workbook, top=0, left=0, right=0, bottom=2))
                    worksheet.set_column('C:C', 18)
                    worksheet.write(index_row-2, 3, "", headers.get_cell_format(workbook, top=0, left=0, right=0, bottom=2, liquid_de=contractor))

                    _range = 'A' + str(index_row-1) + ':C' + str(index_row-1)
                    worksheet.merge_range(_range, "")

                    for iter, row in enumerate(df_cargas_totales.to_frame().iterrows()):
                        total_combustible += row[1][0]
                        worksheet.write(index_row-2, 3, row[1][0], headers.get_cell_format(workbook, top=0, left=0, right=0, bottom=2, liquid_de=contractor))

                if total_combustible > 0.0:
                    index_row_descuento_combustible_o_total_a_cobrar = write_total_fuel_consumption(writer,
                                                                                                sheetName,
                                                                                                index_row_descuento_combustible_o_total_a_cobrar,
                                                                                                workbook, worksheet, total_combustible * -1)

            if total_combustible > 0.0:
                write_total_to_be_charged(writer, sheetName, index_row_descuento_combustible_o_total_a_cobrar + 1 , workbook, worksheet, total_a_cobrar - total_combustible)
            else:
                write_total_to_be_charged(writer, sheetName, index_row_descuento_combustible_o_total_a_cobrar, workbook, worksheet, total_a_cobrar)

        data = {'EMPRESA':  [transportista], 'BANCO':[banco], 'TIPO CUENTA':[tipo_cuenta], 'CUENTA':[cuenta], 'ANTERIOR':[anterior], 'IMPORTE A GIRAR':[total_a_cobrar - total_combustible], 'RUT':[rut], 'RETENCION':[retencion_iva], 'CERT':[certificado_de_credito], 'DESC. COMBUSTIBLE':[total_combustible] }
        if df_resumen_de_pagos is None:
            df_resumen_de_pagos = pd.DataFrame (data, columns = ['EMPRESA', 'BANCO', 'TIPO CUENTA', 'CUENTA', 'ANTERIOR', 'IMPORTE A GIRAR', 'RUT', 'RETENCION', 'CERT', 'DESC. COMBUSTIBLE'])
        else:
            df_new_resumen_pago = pd.DataFrame (data, columns = ['EMPRESA', 'BANCO', 'TIPO CUENTA', 'CUENTA', 'ANTERIOR', 'IMPORTE A GIRAR', 'RUT', 'RETENCION', 'CERT', 'DESC. COMBUSTIBLE'])
            #df_resumen_de_pagos = df_resumen_de_pagos.append(df_new_resumen_pago, ignore_index=True)
            df_resumen_de_pagos = pd.concat([df_resumen_de_pagos, df_new_resumen_pago], ignore_index=True)
        return df_resumen_de_pagos
