import pandas as pd
import random
import numpy as np
import matplotlib.pyplot as plt
import re
from matplotlib.backends.backend_pdf import PdfPages
import xlwings as xw
from plotly import express as px
from fpdf import FPDF
import dataframe_image as dfi

pedidos_sucio = pd.read_csv('order_details.csv', sep=';')


def extract():
    df_precio = pd.read_csv('pizzas.csv', sep=',')
    pedidos = pd.read_csv('order_details.csv', sep=';')
    pizza_ingrediente = pd.read_csv('pizza_types.csv', encoding='latin-1')
    pedidos_fechas = pd.read_csv('orders.csv', sep=';')

    return pedidos, pizza_ingrediente, df_precio


def transform(pedidos, pizza_ingrediente, df_precio):
    # Transformation of the quantity

    pedidos['quantity'] = pedidos['quantity'].fillna(
        1)  # fill the NaN values with 1
    pedidos['quantity'] = pedidos['quantity'].replace('One', 1)
    pedidos['quantity'] = pedidos['quantity'].replace('one', 1)
    pedidos['quantity'] = pedidos['quantity'].replace('two', 2)
    pedidos['quantity'] = pedidos['quantity'].replace('-1', 1)
    pedidos['quantity'] = pedidos['quantity'].replace('-2', 2)
    # change the type of the column quantity to int
    pedidos['quantity'] = pedidos['quantity'].astype(int)

    pedidos['quantity'].isna().sum()  # check if there are NaN values

    # We want to change the csv to match a pettern in pizza_id and quantity columns
    pedidos['pizza_id'] = pedidos['pizza_id'].str.replace('-', '_')
    pedidos['pizza_id'] = pedidos['pizza_id'].str.replace('@', 'a')
    pedidos['pizza_id'] = pedidos['pizza_id'].str.replace('3', 'e')
    pedidos['pizza_id'] = pedidos['pizza_id'].str.replace('0', 'o')
    pedidos['pizza_id'] = pedidos['pizza_id'].str.replace(' ', '_')

    # Top 5 most common pizza types, we would have do that but the recomendation will be different because of the random function, but
    # it will also be a good estimation
    a = pedidos['pizza_id'].value_counts().head(1)
    # replace the Nan with the top 5 most common pizza and size randomly
    a = a.index
    a = list(a)
    pedidos['pizza_id'].fillna(random.choice(a), inplace=True)

    # We apply the multiplication of the quantity to get an idea of the total amount of pizzas sold
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

    del(pedidos['order_details_id'])
    del(pedidos['order_id'])

    # Merge the two dataframes
    df_rev = df_precio.merge(pedidos, on='pizza_id', how='left')
    # fill the NaN values with the mean of the column
    df_rev['quantity'].fillna(df_rev['quantity'].mean(), inplace=True)
    # calculate the revenue
    df_rev['revenue'] = round(df_rev['quantity']*df_rev['price'], 2)
    # We delete the columns that we don't need anymore
    del(df_rev['size'])
    del(df_rev['price'])

    df_rev = df_rev.groupby('pizza_type_id').sum().sort_values(
        by='quantity', ascending=False)  # group by pizza type and sort by quantity
    df_rev = df_rev.reset_index()

    # We merge the two dataframes to estimate the ingridients
    df_semifinal = pd.merge(df_rev, pizza_ingrediente, on='pizza_type_id').sort_values(
        by='quantity', ascending=False)

    # We calculate the total amount of each ingredient
    diccionario_ingredientes = {}
    for i in range(len(df_semifinal['ingredients'])):
        ingredientes = df_semifinal['ingredients'][i].split(',')
        for ingrediente in ingredientes:
            if ingrediente in diccionario_ingredientes:
                diccionario_ingredientes[ingrediente] += df_semifinal['quantity'][i]
            else:
                diccionario_ingredientes[ingrediente] = df_semifinal['quantity'][i]
    for i in diccionario_ingredientes:
        diccionario_ingredientes[i] = diccionario_ingredientes[i]/52
        diccionario_ingredientes[i] = round(diccionario_ingredientes[i]+1, 0)

    diccionario = {'Ingredientes': list(diccionario_ingredientes.keys(
    )), 'Cantidad': list(diccionario_ingredientes.values())}
    final_df = pd.DataFrame.from_dict(diccionario)

    # We create a new dataframe to see the tipology of the csv

    return final_df, df_rev


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


def load(final_df, df_rev, pedidos_sucio):

    # Parte 2.2 pedidos semanal de ingredientes a csv
    final_df.to_csv('ingridients_per_week.csv',
                    index=False, header=True, sep='.')

    # Bloque 3.2
    # To xml
    final_df.to_xml('lista_ingredientes.xml')

    wb = xw.Book()
    sht = wb.sheets[0]
    sht.name = 'Reporte pizzas por más ingresos'

    # Hoja 1
    # Pie chart pizzas por revenue
    insert_heading(sht.range('A2'), 'Pizzas más vendidas por ingresos')

    values = df_rev['revenue']
    names = df_rev['pizza_type_id']
    fig1 = px.pie(df_rev,
                  values=values,
                  names=names,
                  title='Pizzas más vendidas por ingresos')
    fig1.update_traces(textposition='inside', textinfo='percent+label')
    sht.pictures.add(fig1, name='PieChart1', update=True,
                     left=sht.range('A5').left, top=sht.range('A5').top, width=500, height=400)

    # fig1.write_image(
    #     "C:/Users/pablo/OneDrive/Escritorio/2º iMat/Adquisición de Datos/Bloque2/Maven Pizzas 2016/pie_chart.png")
    sht1 = wb.sheets.add('Ingredientes por semana')
    # Hoja 2
    # Gráfica de barras de ingredientes
    sht1.name = 'Reporte de ingredientes '
    insert_heading(sht1.range('A2'), 'Ingredientes por semana')

    fig0 = plt.figure()
    x = final_df['Ingredientes']
    y = final_df['Cantidad']
    ax = plt.subplot()
    plt.bar(x, y)
    plt.setp(ax.get_xticklabels(), rotation=90, horizontalalignment='right')
    plt.grid(False)
    plt.title('Ingredientes por semana')
    sht1.pictures.add(fig0,  name='MyPlot', update=True, left=sht1.range(
        'A4').left, top=sht1.range('A4').top, width=450, height=500)
    fig0.savefig(
        'C:/Users/pablo/OneDrive/Escritorio/2º iMat/Adquisición de Datos/Bloque2/Maven Pizzas 2016/barras_ingredientes.png', bbox_inches='tight')
    # Gráfica de barras de pizzas
    sht2 = wb.sheets.add()
    sht2.name = 'Reporte de pizzas'
    insert_heading(sht2.range('A2'), 'Pizzas anuales')
    fig = plt.figure()
    x = df_rev['pizza_type_id']
    y = df_rev['quantity']
    plt.bar(x, y)
    plt.grid(False)
    plt.title('Pizzas anuales')
    plt.xticks(rotation=90)
    sht2.pictures.add(fig,  name='MyPlot1', update=True, left=sht2.range(
        'A4').left, top=sht2.range('A4').top, width=450, height=500)
    # fig.savefig(
    #     'C:/Users/pablo/OneDrive/Escritorio/2º iMat/Adquisición de Datos/Bloque2/Maven Pizzas 2016/barras_pizzas.png', bbox_inches='tight')
# PEDIDO SEMANAL
    sht1.range('J4').value = final_df
    insert_heading(sht1.range('J2'), 'Pedido semanal')

    # Informe de calidad de datos
    # Type of each column in the dataframe
    sht3 = wb.sheets.add()
    sht3.name = 'Informe de calidad de datos'
    insert_heading(sht3.range('A2'), 'Calidad de datos de los dataframe')
    df_calidad_final = pd.DataFrame({'Nulls': final_df.isnull().sum(),
                                     'NaNs': final_df.isna().sum()
                                     })

    df_calidad_orders = pd.DataFrame({'Nulls': pedidos_sucio.isnull().sum(),
                                      'NaNs': pedidos_sucio.isna().sum()
                                      })

    df_calidad_pizza_ingrediente = pd.DataFrame({'Nulls': pizza_ingrediente.isnull().sum(),
                                                 'NaNs': pizza_ingrediente.isna().sum()})

    df_calidad_revenue = pd.DataFrame({'Nulls': df_rev.isnull().sum(),
                                       'NaNs': df_rev.isna().sum()})

    insert_subheading(sht3.range('A4'),
                      'Calidad de datos del dataframe final')
    sht3.range('A6').value = df_calidad_final
    insert_subheading(sht3.range('A13'),
                      'Calidad de datos del dataframe pedidos')
    sht3.range('A15').value = df_calidad_orders
    insert_subheading(sht3.range('E4'),
                      'Calidad de datos del dataframe pizza_ingrediente')
    sht3.range('E6').value = df_calidad_pizza_ingrediente

    insert_subheading(sht3.range('E13'),
                      'Calidad de datos del dataframe revenue')
    sht3.range('E15').value = df_calidad_revenue
    wb.save('ingridients_per_week.xlsx')

    # Bloque 4.1
    pdf = FPDF()
    pdf.add_page()
    # Add graphs
    pdf.set_font("Arial", size=18)
    pdf.cell(200, 10, txt="Maven Pizzas Report of 2016",
             ln=5, align='C')

    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, txt="Analysis of the ingredients used in the pizzas, and sales income",
             ln=50, align='C')
    pdf.image('Portada_maven.PNG', w=180, h=195)

    pdf.add_page()
    pdf.set_font("Arial", size=18)
    pdf.cell(200, 10, txt="Pizzas most sold by income",
             ln=5, align='C')
    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, txt="Como podemos observar la pizza que más dinero genera es la big meat, seguida de la five cheese y ",
             ln=15, align='L')
    pdf.cell(200, 10, txt="la thai chicken",
             ln=15, align='L')
    pdf.image('pie_chart.png', w=180, h=195)

    pdf.add_page()
    pdf.set_font("Arial", size=18)
    pdf.cell(200, 10, txt="Ingredients used per week are shown in the following graph",
             ln=5, align='C')
    pdf.image('barras_ingredientes.png', w=180, h=195)
    # Add a df to the pdf
    dfi.export(final_df, 'df.png', max_rows=73, max_cols=3,
               table_conversion='matplotlib')
    pdf.add_page()
    pdf.set_font("Arial", size=14)
    pdf.cell(200, 10, txt="The following table shows the ingredients used per week",
             ln=5, align='C')
    pdf.image('df.png', w=130, h=245)

    pdf.add_page()
    pdf.set_font("Arial", size=18)
    pdf.cell(200, 10, txt=" Pizzas sold per year",
             ln=5, align='C')

    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, txt=" Como podemos observar estas son las pizzas más vendidas en el año 2016, en las que casualmente",
             ln=5, align='L')
    pdf.cell(200, 10, txt=" se encuentran los ingredientes más utilizados en las pizzas, por lo que podemos concluir que",
             ln=5, align='L')
    pdf.cell(200, 10, txt=" los ingredientes más utilizados se encuentran en las pizzas más vendidas y en las que más generan",
             ln=5, align='L')
    pdf.image('barras_pizzas.png', w=180, h=195)
    pdf.output('maven_pizzas_2016.pdf', 'F')

    # return final_df, df_rev
if __name__ == '__main__':
    pedidos, pizza_ingrediente, df_precio = extract()
    final_df, df_rev = transform(pedidos, pizza_ingrediente, df_precio)
    load(final_df, df_rev, pedidos_sucio)
