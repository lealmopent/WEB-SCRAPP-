from bs4 import BeautifulSoup
import requests
import sys
import urllib, urllib.request
import re
import xlwt
import time
z= (".txt")
print("\n---Los Angeles Dodgers---")
#Punto 3 - Definir nombre del archivo.
x= str(input("Ingrese el nombre del archivo: "))
y= (x + z)
arch= open(y, "w") 
#Punto 1 - Definir las páginas web en las que se hará la investigación .
#Punto 2 - Establecer paginas donde se obtendra la informacion.
arch.write("https://www.mlb.com/es/dodgers\nhttps://twitter.com/Dodgers\nhttps://www.facebook.com/Dodgers/\nhttps://www.instagram.com/dodgers/?hl=es-la")
arch.close() 
arch=open(y, "r") 
mensaje= arch.read() 
arch.close() 
#Punto 4 - Leer, agregar, actualizar y modificar los enlaces almacenados.   
def Menu():
    print("---------Menu---------\n1) Leer\n2) Agregar\n3) Actualizar\n4) Modificar\n5) Salir\n-----------------------")
    x= int(input("Ingresa el numero: "))
    print("-----------------------")
    if x==1:
        print(mensaje)
    elif x==2:
        arch= open(y,"a")
        url= str(input("Introduce el URL: "))
        direccion =re.compile(r'https://www') 
        a =direccion.search(url) 
        if a is None: 
            print ("¡El URL es incorrecto!")
        else: 
            print("¡El URL es correcto!")
            arch.write(url + "\n")
            arch.close()
    elif x==3:
        arch= open(y, "r")
        arch.seek(0)
        print(arch.read())
        arch.close()
    elif x==4:
        arch= open(y, "r")
        arch.seek(0)
        print(arch.read())
        arch.close()
        arch= open(y,"r+")
        archivo= arch.readlines()
        print("-----------------------\n")
        num= int(input("¿Que linea desea cambiar?\nIngrese numero de linea: "))
        b= (archivo[num-1])
        ruta= str(input("Introduce el URL: "))
        direccion =re.compile(r'https://www') 
        aa =direccion.search(ruta) 
        if aa is None: 
            print ("¡El URL es incorrecto!")
            Menu()
        else: 
            print("¡El URL es correcto!")
            arch.seek(0)
            arch.writelines(b + ruta)
            arch.close()
    elif x==5:
        print("----> Exit")
    elif x>5:
        print("¡Ingresa valor valido!")
        Menu()
    def Minimenu():
        print("----------------------------------\n¿Desea realizar otra modificacion?\n1)Si\n2)No\n----------------------------------")
        m= int(input("Ingresa el nuemro: "))
        if m==1:
            Menu()
        elif m==2:
            print("----> Exit")
            print("----------------------------------")
        elif m>2:
            print("¡Ingresa valor valido!")
            Minimenu()
    Minimenu()
Menu()
#Punto 5 - Hacer un request para guardar las páginas a consultar.
print("Realizando Request...")
lista=['https://www.mlb.com/es/dodgers', 'https://www.facebook.com/Dodgers/','https://twitter.com/Dodgers', 'https://www.instagram.com/dodgers/?hl=es-la']
fin=len(lista)
for i in range(0,fin):
	url=lista[i]
	st='Pagina'
	ext='.txt'
	num=str(i)
	pag=(st+num+ext)
	ruta=("C:/Users/carlo/Desktop/Python On/PIA/"+pag)
	with urllib.request.urlopen(url) as html_handler:
		html=html_handler.read()
		with open(ruta, 'wb') as file_handler:
			print("Escribiendo pagina... ", i)
			file_handler.write(html)
print("¡Request Exitoso!")
print("----------------------------------")
#Punto 6 - Obtener información significativa de los jugadores.
print("\n-----Bienvenido al buscador-----\n")
baseurl='https://www.mlb.com/player/'
print("-----Lista de Jugadores-----\n")
print("--Nombre del jugador--    ---Codigo---")
print("Scott Fernandez  #75: scott-alexander-518397\nPedro Báez #52: pedro-baez-520980\nWalker Buehler #21: walker-buehler-621111\nCaleb Ferguson #64: caleb-ferguson-657571\nDylan Floro51: dylan-floro-571670\nTony Gonsolin #46: tony-gonsolin-664062\nBrusdar Graterol #48: brusdar-graterol-660813\nKenley Jansen #74: kenley-jansen-445276\nJoe Kelly #17: joe-kelly-523260\nClayton Kershaw #22: clayton-kershaw-477132\nAdam Kolarek #56: adam-kolarek-592473\nDustin May #85: dustin-may-669160\nJimmy Nelson #40: jimmy-nelson-51907\nDavid Price #33: david-price-456034\nDennis Santana #77: dennis-santana-642701\nRoss Stripling #68: ross-stripling-548389\nBlake Treinen #49: blake-treinen-595014\nJulio Urías #7: julio-urias-628711\nAlex Wood #57: alex-wood-622072\nAustin Barnes #15: austin-barnes-605131\nWill Snith #16: will-smith-669257\nMatt Beaty #45: matt-beaty-607461\nEnrique Hernández #14: enrique-hernandez-571771\nGavin Lux #9: gavin-lux-666158\nMax Muncy #13: max-muncy-571970\nEdwin Rios #43: edwin-rios-621458\nCorey Seager #5: corey-seager-608369\nJustin Turner #10: justin-turner-457759\nCody Bellinger #35: cody-bellinger-641355\nMookie Betts #50: mookie-betts-605141\nJoc Pederson #31: joc-pederson-592626\nA.J Pollock #11: a-j-pollock-572041\nChris Taylor #3: chris-taylor-621035\n")
jugador= str(input("Introduzca el codigo del jugador que desea conocer: "))
page = requests.get(baseurl+jugador)
print(page.status_code)
soup = BeautifulSoup(page.content ,"html.parser")
inf = soup.find_all("ul")
infText= inf[1].getText()
print (infText)
def Jugadores():
        print("-----------------------------")
        replay = input("¿Desea continuar la busqueda?\n1)Si\n2)No\n-----> ").lower()
        if replay in ("1"):
            jugador = str(input("Introduzca un jugador: "))
            print("-----------------------------")
            page = requests.get(baseurl+jugador)
            print(page.status_code)
            soup = BeautifulSoup(page.content ,"html.parser")
            inf = soup.find_all("ul")
            infText= inf[1].getText()
            print (infText)
            Jugadores()
        else: 
            replay in ("2")
            print("Saliendo de programa...")
Jugadores()
#Punto 7 - Definir una gira para el equipo de beisbol.
#GIRA OFICIAL
arch=open("Gira2020.txt","w")  
arch.write('\n----------Gira Junio 2020----------\n\nViernes 5 de junio: Los Ángeles Dodgers vs Colorado Rockies\nHora: 7:15 pm\nRecinto: Dodger Stadium\n\nSabado 6 de junio: Los Ángeles Dodgers vs New York Mets\nHora: 7:15 pm\nRecinto: Citi Field\n\nDomingo 7 de junio: Los Ángeles Dodgers vs Minnesota Twins\nHora: 8:00 pm\nRecinto: Target Field\n\nLunes 8 de junio: Los Ángeles Dodgers vs Arizona Diamondbacks\nHora: 6:45 pm\nRecinto: Chase Field\n\nMartes 9 de junio: Los Ángeles Dodgers vs Philadelphia Phillies\nHora: 8:15 pm\nRecinto: Citizens Bank Park\n\nMiercoles 10 de junio: Los Ángeles Dodgers vs San Diego Padres\nHora: 6:15 pm\nRecinto: Petco Park\n\nJueves 11 de junio: Los Ángeles Dodgers vs Cincinnati Reds\nHora: 9:00 pm\nRecinto: Progressive Field\n\nViernes 12 de junio: Los Ángeles Dodgers vs Toronto Blue Jays\nHora: 8:00 pm\nRecinto: Rogers Centre\n\nSabado 13 de junio: Los Ángeles Dodgers vs Kansas City Royals\nHora: 7:30 pm\nRecinto: Kauffman Stadium\n\nDomingo 14 de junio: Los Ángeles Dodgers vs Pittsburgh Pirates\nHora: 5:00 pm\nRecinto: PNC Park')
arch.close() 
arch=open('Gira2020.txt','r') 
mensaje=arch.read() 
print(mensaje) 
arch.close()
#CLIMA OFICIAL
url=("http://api.openweathermap.org/data/2.5/weather?id=5344994&appid=2f2e398e6ba67f721f728b7d94a11bd4")
datos= requests.get(url)
read= datos.json()
print(("Nombre de la ciudad: {}".format(read["name"])))
print(("Temperatura: {}".format(read["main"]["temp"])))
print(("Humedad: {}".format(read["main"]["humidity"])))
print(("Presion: {}".format(read["main"]["pressure"])))
print(("Viento: {}".format(read["wind"]["speed"])))
print(("Descripcion: {}\n".format(read["weather"][0]["description"])))

url=("http://api.openweathermap.org/data/2.5/weather?id=5128638&appid=2f2e398e6ba67f721f728b7d94a11bd4")
datos= requests.get(url)
read= datos.json()
print(("Nombre de la ciudad: {}".format(read["name"])))
print(("Temperatura: {}".format(read["main"]["temp"])))
print(("Humedad: {}".format(read["main"]["humidity"])))
print(("Presion: {}".format(read["main"]["pressure"])))
print(("Viento: {}".format(read["wind"]["speed"])))
print(("Descripcion: {}\n".format(read["weather"][0]["description"])))

url=("http://api.openweathermap.org/data/2.5/weather?id=5037779&appid=2f2e398e6ba67f721f728b7d94a11bd4")
datos= requests.get(url)
read= datos.json()
print(("Nombre de la ciudad: {}".format(read["name"])))
print(("Temperatura: {}".format(read["main"]["temp"])))
print(("Humedad: {}".format(read["main"]["humidity"])))
print(("Presion: {}".format(read["main"]["pressure"])))
print(("Viento: {}".format(read["wind"]["speed"])))
print(("Descripcion: {}\n".format(read["weather"][0]["description"])))

url=("http://api.openweathermap.org/data/2.5/weather?id=5551665&appid=2f2e398e6ba67f721f728b7d94a11bd4")
datos= requests.get(url)
read= datos.json()
print(("Nombre de la ciudad: {}".format(read["name"])))
print(("Temperatura: {}".format(read["main"]["temp"])))
print(("Humedad: {}".format(read["main"]["humidity"])))
print(("Presion: {}".format(read["main"]["pressure"])))
print(("Viento: {}".format(read["wind"]["speed"])))
print(("Descripcion: {}\n".format(read["weather"][0]["description"])))

url=("http://api.openweathermap.org/data/2.5/weather?id=5205788&appid=2f2e398e6ba67f721f728b7d94a11bd4")
datos= requests.get(url)
read= datos.json()
print(("Nombre de la ciudad: {}".format(read["name"])))
print(("Temperatura: {}".format(read["main"]["temp"])))
print(("Humedad: {}".format(read["main"]["humidity"])))
print(("Presion: {}".format(read["main"]["pressure"])))
print(("Viento: {}".format(read["wind"]["speed"])))
print(("Descripcion: {}\n".format(read["weather"][0]["description"])))

url=("http://api.openweathermap.org/data/2.5/weather?id=4379545&appid=2f2e398e6ba67f721f728b7d94a11bd4")
datos= requests.get(url)
read= datos.json()
print(("Nombre de la ciudad: {}".format(read["name"])))
print(("Temperatura: {}".format(read["main"]["temp"])))
print(("Humedad: {}".format(read["main"]["humidity"])))
print(("Presion: {}".format(read["main"]["pressure"])))
print(("Viento: {}".format(read["wind"]["speed"])))
print(("Descripcion: {}\n".format(read["weather"][0]["description"])))

url=("http://api.openweathermap.org/data/2.5/weather?id=5165418&appid=2f2e398e6ba67f721f728b7d94a11bd4")
datos= requests.get(url)
read= datos.json()
print(("Nombre de la ciudad: {}".format(read["name"])))
print(("Temperatura: {}".format(read["main"]["temp"])))
print(("Humedad: {}".format(read["main"]["humidity"])))
print(("Presion: {}".format(read["main"]["pressure"])))
print(("Viento: {}".format(read["wind"]["speed"])))
print(("Descripcion: {}\n".format(read["weather"][0]["description"])))

url=("http://api.openweathermap.org/data/2.5/weather?id=5174095&appid=2f2e398e6ba67f721f728b7d94a11bd4")
datos= requests.get(url)
read= datos.json()
print(("Nombre de la ciudad: {}".format(read["name"])))
print(("Temperatura: {}".format(read["main"]["temp"])))
print(("Humedad: {}".format(read["main"]["humidity"])))
print(("Presion: {}".format(read["main"]["pressure"])))
print(("Viento: {}".format(read["wind"]["speed"])))
print(("Descripcion: {}\n".format(read["weather"][0]["description"])))

url=("http://api.openweathermap.org/data/2.5/weather?id=4711801&appid=2f2e398e6ba67f721f728b7d94a11bd4")
datos= requests.get(url)
read= datos.json()
print(("Nombre de la ciudad: {}".format(read["name"])))
print(("Temperatura: {}".format(read["main"]["temp"])))
print(("Humedad: {}".format(read["main"]["humidity"])))
print(("Presion: {}".format(read["main"]["pressure"])))
print(("Viento: {}".format(read["wind"]["speed"])))
print(("Descripcion: {}\n".format(read["weather"][0]["description"])))

url=("http://api.openweathermap.org/data/2.5/weather?id=6254927&appid=2f2e398e6ba67f721f728b7d94a11bd4")
datos= requests.get(url)
read= datos.json()
print(("Nombre de la ciudad: {}".format(read["name"])))
print(("Temperatura: {}".format(read["main"]["temp"])))
print(("Humedad: {}".format(read["main"]["humidity"])))
print(("Presion: {}".format(read["main"]["pressure"])))
print(("Viento: {}".format(read["wind"]["speed"])))
print(("Descripcion: {}\n".format(read["weather"][0]["description"])))
#CALENDARIO OFICIAL
arch=open("Calendario2020.txt","w")  
arch.write('----------Calendario Junio 2020----------\n\nViernes 5 de junio: Los Ángeles Dodgers vs Colorado Rockies\nHora: 7:15 pm\nRecinto: Dodger Stadium\nNombre de la ciudad: East Los Angeles\nTemperatura: 294.49\nHumedad: 60\nPresion: 1009\nViento: 1.5\nDescripcion: clear sky\n\nSabado 6 de junio: Los Ángeles Dodgers vs Nueva York Mets\nHora: 7:15 pm\nRecinto: Citi Field\nNombre de la ciudad: New York\nTemperatura: 295.5\nHumedad: 53\nPresion: 1010\nViento: 3.1\nDescripcion: overcast clouds\n\nDomingo 7 de junio: Los Ángeles Dodgers vs Minnesota Twins\nHora: 8:00 pm\nRecinto: Target Field\nNombre de la ciudad: Minnesota\nTemperatura: 294.73\nHumedad: 49\nPresion: 1008\nViento: 2\nDescripcion: clear sky\n\nLunes 8 de junio: Los Ángeles Dodgers vs Arizona Diamondbacks\nHora: 6:45 pm\nRecinto: Chase Field\nNombre de la ciudad: Arizona City\nTemperatura: 314.31\nHumedad: 11\nPresion: 1004vViento: 1.5\nDescripcion: clear sky\n\nMartes 9 de junio: Los Ángeles Dodgers vs Philadelphia Phillies\nHora: 8:15 pm\nRecinto: Citizens Bank Park\nNombre de la ciudad: Philadelphia\nTemperatura: 298.35\nHumedad: 66\nPresion: 1011\nViento: 3.29\nDescripcion: moderate rain\n\nMiercoles 10 de junio: Los Ángeles Dodgers vs San Diego Padres\nHora: 6:15 pm\nRecinto: Petco Park\nNombre de la ciudad: California\nTemperatura: 299.2\nHumedad: 78\nPresion: 1009\nViento: 2.6\nDescripcion: clear sky\n\nJueves 11 de junio: Los Ángeles Dodgers vs Cincinnati Reds\nHora: 9:00 pm\nRecinto: Progressive Field\nNombre de la ciudad: Ohio\nTemperatura: 294.53\nHumedad: 88\nPresion: 1011\nViento: 1.5\nDescripcion: clear sky\n\nViernes 12 de junio: Los Ángeles Dodgers vs Toronto Blue Jays\nHora: 8:00 pm\nRecinto: Rogers Centre\nNombre de la ciudad: Toronto\nTemperatura: 290.59\nHumedad: 87\nPresion: 1010\nViento: 1.38\nDescripcion: overcast clouds\n\nSabado 13 de junio: Los Ángeles Dodgers vs Kansas City Royals\nHora: 7:30 pm\nRecinto: Kauffman Stadium\nNombre de la ciudad: Missouri City\nTemperatura: 301.14\nHumedad: 78\nPresion: 1011\nViento: 6.2\nDescripcion: scattered clouds\n\nDomingo 14 de junio: Los Ángeles Dodgers vs Pittsburgh Pirates\nHora: 5:00 pm\nRecinto: PNC Park\nNombre de la ciudad: Pennsylvania\nTemperatura: 293.1\nHumedad: 88\nPresion: 1013\nViento: 2.6\nDescripcion: heavy intensity rain')
arch.close() 
arch=open('Calendario2020.txt','r') 
mensaje=arch.read() 
print(mensaje) 
arch.close()
print("----------------------------------------------------")
print("          Calendario de la gira completado          ")
print("----------------------------------------------------")
#Punto 9 - Documento de Excel 
time.sleep(2)
print("Generando Archivo Excel...")
lista=['https://www.mlb.com/es/dodgers','https://es.wikipedia.org/wiki/Los_Angeles_Dodgers','https://www.instagram.com/dodgers/?hl=es-la','https://www.facebook.com/Dodgers/','https://www.tiktok.com/@dodgers','https://twitter.com/DodgersNation']
lista2=['Pagina Web','wikipedia','Instagram', 'facebook', 'tiktok', 'twitter']
nombres=['Scott Alain Alexander','Pedro Alberys Báez', 'Walker Anthony Buehler', 'Caleb Paul Ferguson', 'Dylan Lee Floro', 'Anthony D. Gonsolin', 'Brusdar Javier Graterol', 'Kenley Geronimo Jansen', 'Joseph William Kelly', 'Adam John Kolarek', 'Dustin Jake May', 'David Taylor Price', 'Dennis Anfernee Santana', 'Thomas Ross Stripling', 'Blake M. Treinen', 'Julio Cesar Urías', 'Robert Alexander Wood', 'Austin Scott Barnes', 'William Dills Smith', 'Matthew Thomas Beaty', 'Enrique J. Hernández', 'Fullname: Gavin Thomas Lux', 'Maxwell Steven Muncy', 'Edwin Gabriel Ríos', 'Corey Drew Seager', 'Justin Matthew Turner', 'Cody James Bellinger', 'Markus Lynn Betts', 'Joc R. Pederson', 'Allen Lorenz Pollock', 'Christopher Armand Taylor']
nac=['7/10/1989','3/11/1988', '7/28/1994', ' 7/02/1996', '12/27/1990', '5/14/1994', '8/26/1998', '9/30/1987', '6/09/1988', '1/14/1989', '9/06/1997', '8/26/1985', '4/12/1996', '11/23/1989', '6/30/1988', '8/12/1996', '1/12/1991', '12/28/1989', '3/28/1995', '4/28/1993', '8/24/1991', '11/23/1997', '8/25/1990', '4/21/1994', '4/27/1994', '11/23/1984', '7/13/1995', '10/07/1992', '4/21/1992', '12/05/1987', '8/29/1990']
lug=['Santa Rosa, CA', 'Bani, Dominican Republic', 'Lexington, KY', 'Columbus, OH', 'Merced, CA', 'Vacaville, CA', 'Calobozo, Venezuela', 'Willemstad, Curacao', 'Anaheim, CA', 'Baltimore, MD',' Justin, TX', 'Murfreesboro, TN', 'Dominican Republic', 'Bluebell, PA', 'Kansas', 'Culiacan Rosales, Sinaloa, Mexico', 'Charlotte, NC', 'Riverside, CA', 'Louisville, KY', 'Snellville, GA', 'San Juan, Puerto Rico', 'Kenosha, WI', 'Midland, TX', 'Caguas, Puerto Rico', 'Charlotte, NC', 'Long Beach, CA', 'Scottsdale, AZ', 'Nashville, TN', 'Palo Alto, CA', 'Hebron, CT', 'Virginia Beach, VA']
redesn=['Kenley Geronimo Jansen', 'Anthony D. Gonsolin', 'Kenley Geronimo Jansen', 'Dustin Jake May', 'David Taylor Price', 'Dennis Anfernee Santana', 'Thomas Ross Stripling', 'Robert Alexander Wood', 'Maxwell Steven Muncy', 'Edwin Gabriel Ríos', 'Corey Drew Seager', 'Justin Matthew Turner']
redeslin=['https://twitter.com/kenleyjansen74', 'https://twitter.com/goooose15', 'https://twitter.com/kenleyjansen74', 'https://twitter.com/d_maydabeast', 'https://twitter.com/DAVIDprice24', 'https://twitter.com/RossStripling', 'https://twitter.com/Awood45', 'https://twitter.com/maxmuncy9', 'https://twitter.com/Edwin_Rios30', 'https://twitter.com/CoreyDrewSeager', 'https://twitter.com/redturn2']
dia=['Viernes 5 de junio','Sabado 6 de junio','Domingo 7 de junio', 'Lunes 8 de junio', 'Martes 9 de junio', 'Miercoles 10 de junio', 'Jueves 11 de junio', 'Viernes 12 de junio','Sabado 13 de junio','Domingo 14 de junio']
hora=['7:15 pm', '7:15 pm', '8:00 pm', '6:45 pm', '8:15 pm', '6:15 pm', '9:00 pm','8:00 pm', '7:30 pm', '5:00 pm']
estadio=['Dodger Stadium','Citi Field','Target Field','Chase Field','Citizens Bank Park','Petco Park','Progressive Field','Rogers Centre','Kauffman Stadium','PNC Park']
ciudad=['East Los Angeles','New York','Minnesota','Arizona City','Philadelphia','California','Ohio','Toronto','Missouri City','Pennsylvania']
tem=['294.49','295.5','294.73','314.31','298.35','299.2','294.53','290.59','301.14','293.1']
hum=['60','53','49','11','66','78','88','87','78','88']
presion=['1009','1010','1008','1004','1011','1009','1011','1010','1011','1013']
viento=['1.5','3.1','2','1.5','3.29','2.6','1.5','1.38','6.2','2.6']
des=['clear sky','overcast clouds','clear sky','clear sky','moderate rain','clear sky','clear sky','overcast clouds','scattered clouds','heavy intensity rain']
book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet("Sheet 1")
sheet1.write(4, 0, "Nombre de red social")
sheet1.write(4, 1, "URL")
sheet1.write(12, 0, "Nombre Completo del  Jugador")
sheet1.write(12, 1, "Fecha de Nacimiento")
sheet1.write(12, 2, "Lugar de Nacimiento")
sheet1.write(49, 0, "Nombre Jugador")
sheet1.write(49, 1, "Red Social")
sheet1.write(0, 8, "Dia de juego")
sheet1.write(0, 9, "Hora de juego")
sheet1.write(0, 10, "Estadio")
sheet1.write(0, 11, "Ciudad")
sheet1.write(0, 12, "Temperatura")
sheet1.write(0, 13, "Humedad")
sheet1.write(0, 14, "Presion")
sheet1.write(0, 15, "Viento")
sheet1.write(0, 16, "Clima")
i=4
for n in lista:
       i = i+1
       sheet1.write(i, 1, n)
i=4
for n in lista2:
       i = i+1
       sheet1.write(i, 0, n)

i=12
for n in nombres:
       i = i+1
       sheet1.write(i, 0, n)
i=12
for n in nac:
       i = i+1
       sheet1.write(i, 1, n)

i=12
for n in lug:
       i = i+1
       sheet1.write(i, 2, n)

i=50
for n in redesn:
       i = i+1
       sheet1.write(i, 0, n)
i=50
for n in redeslin:
       i = i+1
       sheet1.write(i, 1, n)
i=0
for n in dia:
       i = i+1
       sheet1.write(i, 8, n)

i=0
for n in hora:
       i = i+1
       sheet1.write(i, 9, n)

i=0
for n in estadio:
       i = i+1
       sheet1.write(i, 10, n)

i=0
for n in ciudad:
       i = i+1
       sheet1.write(i, 11, n)

i=0
for n in tem:
       i = i+1
       sheet1.write(i, 12, n)

i=0
for n in hum:
       i = i+1
       sheet1.write(i, 13, n)

i=0
for n in presion:
       i = i+1
       sheet1.write(i, 14, n)
i=0
for n in viento:
       i = i+1
       sheet1.write(i, 15, n)
i=0
for n in des:
       i = i+1
       sheet1.write(i, 16, n)
book.save("Info.xls")
time.sleep(2)
print("¡Archivo Completado!")