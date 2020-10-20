import time
from geopy.geocoders import ArcGIS
import folium
import pandas as pd
import os
import json

# print(os.listdir())
# print(os.path.dirname(__file__))

tilesDiff = ["Stamen Terrain","CartoDB positron"]

arquivoBase = 'Localização.xlsx'

print('Para utilizar esse programa, é nessesário que voce tenha o arquivo "Localização.xlsx" no mesmo diretorio desse programa ')
print('Esse arquivo deve conter uma coluna "Localização" com as informações de localização.')

if arquivoBase in os.listdir():

    dfLocation = pd.read_excel(arquivoBase)

    nom = ArcGIS()

    dfLocation['GEOCODE'] = dfLocation['Localização'].apply(nom.geocode)

    dfLocation['Lat'] = dfLocation['GEOCODE'].apply(lambda x: x.latitude if x != None else None)
    dfLocation['Lon'] = dfLocation['GEOCODE'].apply(lambda x: x.longitude if x != None else None)

    # print(dfLocation)

    dfLocation.to_excel(r'Resultado.xlsx', sheet_name='Your sheet name')
    print('\n     -----------------------------------------------------')
    print('                 Programa Executado com Exito             ')
    print('     -----------------------------------------------------')
else:
    print('\n-----------------------------------------------------')
    print('                 Falha no Programa')
    print('-----------------------------------------------------')
    print('MOTIVO:')
    print('\nArquivo "Localização.xlsx" não se encontra no mesmo diretorio desse programa')
    print('')
    print('Garanta que o arquivo esteja no diretório....   '+os.path.dirname(__file__))
    print('Caso não tenha acesso ao arquivo, crie um arquivo excel com o nome "Localização.xlsx"\nque contenha uma coluna "Localização" com os parametros que deseja rodar.')

print('\nLatitude e Longitude já estão disponível no arquivo "Resultado.xlsx" ')

prompt = input('\nGostaria de plotar os dados no mapa?(s/n)\n').lower()

if prompt == 's':

    # Get the path to this file
    thisFile = os.path.dirname(__file__)

    #Reading excel file that was expelled from findGPS .exe
    df_info = pd.read_excel('Resultado.xlsx')
    isNull = list(df_info["Lat"].isnull())
    qtNull = df_info["Lat"].isnull().sum()
    lat = list(df_info["Lat"])
    lon = list(df_info["Lon"])
    loc = list(df_info["Localização"])

    #created map object
    map = folium.Map([(df_info['Lat'].max()+df_info['Lat'].min())/2, (df_info['Lon'].max()+df_info['Lon'].min())/2],
                     zoom_start=5, tiles=tilesDiff[1])

    #cretaed Feature Group
    fg1 = folium.FeatureGroup(name='Points')

    for latitude, longitude, check, local in zip(lat, lon, isNull, loc):
        if not check:
            fg1.add_child(folium.Marker(location=[latitude, longitude],popup=str(local),
                                    icon=folium.Icon(color='blue',prefix='fa',icon='circle')))


    map.add_child(fg1)

    map.add_child(folium.LayerControl())

    map.save('Plot.html')

    # diferentTiles = ["Stamen Terrain","CartoDB positron"]
    print('\nSua opção foi "s" então o arquivo "plot.html" já está disponível.. ')
    if qtNull != 0:
        print('\n------------------------------- OBS -------------------------------------')
        print('Favor verificar a coluna "Lat" e "Lon" do arquivo "Resultado.xlsx".')
        print('Você possui '+str(qtNull)+' localizações que não foram encontradas pelo Geolocalizador')
        print('O GeoLocalizador deixa as colunas em branco e consequentemente a localiação não foi plotada.')
        print('Pode corrigir essa localização no arquivo "Localização.xlsx" e rodar o programa novamente.')
    print('\npode fechar o programa..')
    time.sleep(60 * 2)
else:
    print('\nSua opção foi "n" então pode fechar o programa..')
    time.sleep(60*2)