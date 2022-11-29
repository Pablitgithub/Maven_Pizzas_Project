import pandas as pd
import random
import numpy as np
import matplotlib.pyplot as plt
import re
from matplotlib.backends.backend_pdf import PdfPages
import xlwings as xw
import xlsxwriter


def extract():
    pedidos = pd.read_csv('order_details.csv')
    pizza_ingrediente = pd.read_csv('pizza_types.csv', encoding='latin-1')

    return pedidos, pizza_ingrediente


def transform(pedidos, pizza_ingrediente):
    # Aplicamos los multiplicadores a las pizzas según su tamaño
    for i in range(len(pedidos['pizza_id'])):
        if pedidos['pizza_id'][i][-1] == 'm':
            pedidos['quantity'][i] = 1.25*pedidos['quantity'][i]
        elif pedidos['pizza_id'][i][-1] == 'l':
            if pedidos['pizza_id'][i][-2] == 'x' and pedidos['pizza_id'][i][-3] == 'x':
                pedidos['quantity'][i] = 2*pedidos['quantity'][i]
            elif pedidos['pizza_id'][i][-2] == 'x' and pedidos['pizza_id'][i][-3] != 'x':
                pedidos['quantity'][i] = 1.75*pedidos['quantity'][i]
            else:
                pedidos['quantity'][i] = 1.5*pedidos['quantity'][i]

    # Agrupas los pedidos por pizza y tamaño, y sumamos el número de pizzas

    pedidos = pedidos.groupby('pizza_id').sum().sort_values(
        by='quantity', ascending=False)
    pedidos = pedidos.reset_index()
    # Eliminamos el tamaño de cada pizza al haber aplicado los multiplicadores y volvemos a agrupar en este caso ya para calcular el número de ingredientes

    pedidos['pizza_id'] = pedidos['pizza_id'].str.replace('_s', '')
    pedidos['pizza_id'] = pedidos['pizza_id'].str.replace('_m', '')
    pedidos['pizza_id'] = pedidos['pizza_id'].str.replace('_l', '')
    pedidos['pizza_id'] = pedidos['pizza_id'].str.replace('_xl', '')
    pedidos['pizza_id'] = pedidos['pizza_id'].str.replace('_xxl', '')
    pedidos = pedidos.groupby('pizza_id').sum().sort_values(
        by='quantity', ascending=False)
    pedidos = pedidos.reset_index()

    # Eliminamos las culumnas que no nos interesan
    del(pedidos['order_details_id'])
    del(pedidos['order_id'])

    # Renombramos la columna para que coincida con el nombre de la columna de la tabla de ingredientes
    pizza_ingrediente.rename(
        columns={'pizza_type_id': 'pizza_id'}, inplace=True)

    # Unimos las dos tablas para obtener el número de ingredientes de cada pizza
    df_semifinal = pd.merge(pedidos, pizza_ingrediente, on='pizza_id').sort_values(
        by='quantity', ascending=False)

    # Recorremos la columna de ingrediente por su separador y contamos el número de ingredientes
    diccionario_ingredientes = {}
    for i in range(len(df_semifinal['ingredients'])):
        ingredientes = df_semifinal['ingredients'][i].split(',')

        for ingrediente in ingredientes:  # Recorremos la lista de ingredientes
            if ingrediente in diccionario_ingredientes:
                # Si el ingrediente ya está en el diccionario, sumamos la cantidad de pizzas que lleva
                diccionario_ingredientes[ingrediente] += df_semifinal['quantity'][i]
            else:
                # Si el ingrediente no está en el diccionario, lo añadimos y le asignamos la cantidad de pizzas que lleva
                diccionario_ingredientes[ingrediente] = df_semifinal['quantity'][i]

    # Dividimos el número de ingredientes entre 52 para obtener el número de ingredientes por semana
    for i in diccionario_ingredientes:
        diccionario_ingredientes[i] = diccionario_ingredientes[i]/52
        # Redondeamos al entero más cercano
        diccionario_ingredientes[i] = round(diccionario_ingredientes[i]+1, 0)

    # Creamos un diccionario con los ingredientes y su cantidad para poder crear el dataframe final
    diccionario = {'Ingredientes': list(diccionario_ingredientes.keys(
    )), 'Cantidad': list(diccionario_ingredientes.values())}  # Creamos un diccionario con los ingredientes y su cantidad para poder crear el dataframe

    final_df = pd.DataFrame.from_dict(diccionario)
    return final_df


def insert_heading(rng, text):
    rng.value = text
    rng.font.bold = True
    rng.font.size = 24
    rng.font.color = (0, 0, 139)


def insert_subheading(rng, text):
    rng.value = text
    rng.font.bold = True
    rng.font.size = 8
    rng.font.color = (0, 0, 139)


def load(final_df):

    # CSV FILE
    final_df.to_csv('ingridients_per_week.csv', index=False, header=True)

    # EXCEL FILE DATA CUALITY
    wb = xw.Book()
    sht = wb.sheets[0]
    sht.name = 'Reporte de ingredientes'


# GRAFICO DE BARRAS PANDAS
    insert_heading(sht.range('A2'), 'Ingredientes por semana')
    fig = plt.figure()
    x = final_df['Ingredientes']
    y = final_df['Cantidad']
    plt.bar(x, y)
    plt.grid(False)
    plt.title('Ingredientes por semana')
    plt.xticks(rotation=90)
    sht.pictures.add(fig,  name='MyPlot', update=True, left=sht.range(
        'A4').left, top=sht.range('A4').top, width=450, height=500)

# PEDIDO SEMANAL
    sht.range('J4').value = final_df
    insert_heading(sht.range('J2'), 'Pedido semanal')

# INFORME DE CALIDAD

    insert_heading(sht.range('A40'), 'Calidad de datos del dataframe')
    df_calidad_final = pd.DataFrame({'Nulls': final_df.isnull().sum(),
                                     'NaNs': final_df.isna().sum()})

    df_calidad_orders = pd.DataFrame({'Nulls': pedidos.isnull().sum(),
                                      'NaNs': pedidos.isna().sum()})

    df_calidad_pizza_ingrediente = pd.DataFrame({'Nulls': pizza_ingrediente.isnull().sum(),
                                                 'NaNs': pizza_ingrediente.isna().sum()})

    insert_subheading(sht.range('B42'), 'Calidad de datos del dataframe final')
    sht.range('B44').value = df_calidad_final
    insert_subheading(sht.range('B49'),
                      'Calidad de datos del dataframe pedidos')
    sht.range('B52').value = df_calidad_orders
    insert_subheading(sht.range('B58'),
                      'Calidad de datos del dataframe pizza_ingrediente')
    sht.range('B60').value = df_calidad_pizza_ingrediente

    wb.save('ingridients_per_week.xlsx')


if __name__ == '__main__':
    pedidos, pizza_ingrediente = extract()
    final_df = transform(pedidos, pizza_ingrediente)
    load(final_df)
