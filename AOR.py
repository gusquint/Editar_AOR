from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import gspread
from datetime import date
from indicadoresmensuales import get_observadores
import time
from decouple import config


def month_num_to_word(num,option):
    if option==1:
        meses=["ene","feb","mar","abr","may","jun","jul","ago","sep","oct","nov","dic"]
    if option==2:
        meses=['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE',"DICIEMBRE"]
    return meses[num-1]


def get_month(word):
    if word.lower()=="today":
        today=date.today()
        month=int(today.strftime("%d/%m/%Y").split("/")[1])
    else:
        month=int(word)
    return month
        

def editar_tarde(month,observadores,lugar):    
    comentario=""
    comentario_counter=""
    counter=0
    if len(observadores)==0:
        comentario_counter="Ningun estudiante tuvo anotacion en el observador por llegadas tarde"
    else:
        for estudiante in observadores:
            for anotacion in observadores[estudiante]:
                if anotacion["codigo"]=="68.1.1":
                    counter+=1
                    comentario+=f"{estudiante} ({anotacion['responsable']}),  "
            comentario_counter=str(counter)+"  "+comentario 

    if comentario_counter=="0  ":
        comentario_counter="Ningun estudiante tuvo anotacion en el observador por llegadas tarde"

    if lugar=="colegio":
        sa=gspread.service_account(filename="C:/Users/gquintero/Desktop/python/pythontest-361720-6e048b13eddb.json")
    else:
        sa=gspread.service_account(filename="C:/Users/PC/Desktop/python/pythontest-361720-6e048b13eddb.json")

    
    google_sheet=sa.open("AOR Gustavo Quintero")
    hoja=google_sheet.worksheet("AÑO")
    col=hoja.find(month)
    hoja.update_cell(22,col.col,comentario_counter)
    print("Se actualizo la celda de llegadas tarde del AOR")
    

def editar_uniforme(month,observadores,lugar):    
    comentario=""
    comentario_counter=""
    counter=0
    if len(observadores)==0:
        comentario_counter="Ningun estudiante tubo anotacion en el observador por uniforme"
    else:
        for estudiante in observadores:
            for anotacion in observadores[estudiante]:
                if anotacion["codigo"]=="68.1.33":
                    counter+=1
                    comentario+=f"{estudiante},  "
            comentario_counter=str(counter)+"  "+comentario 

    if comentario_counter=="0  ":
        comentario_counter="Ningun estudiante tuvo anotacion en el observador por Uniforme"

    if lugar=="colegio":
        sa=gspread.service_account(filename="C:/Users/gquintero/Desktop/python/pythontest-361720-6e048b13eddb.json")
    else:
        sa=gspread.service_account(filename="C:/Users/PC/Desktop/python/pythontest-361720-6e048b13eddb.json")
        
    google_sheet=sa.open("AOR Gustavo Quintero")
    hoja=google_sheet.worksheet("AÑO")
    col=hoja.find(month)
    hoja.update_cell(21,col.col,comentario_counter)
    print("Se actualizo la celda de uniforme del AOR")



def editar_escuela_virtual(month,lugar):    

    if lugar=="colegio":
        sa=gspread.service_account(filename="C:/Users/gquintero/Desktop/python/pythontest-361720-6e048b13eddb.json")
    else:
        sa=gspread.service_account(filename="C:/Users/PC/Desktop/python/pythontest-361720-6e048b13eddb.json")

    google_sheet=sa.open("Escuela virtual 10b 2022")
    hoja=google_sheet.worksheet("Hoja 1")
    col=hoja.find(month)
    num_escuelas=hoja.cell(25,col.col).value
    google_sheet=sa.open("AOR Gustavo Quintero")
    hoja=google_sheet.worksheet("AÑO")
    col=hoja.find(month)
    hoja.update_cell(29,col.col,f"Este mes el {int(num_escuelas)*100//23}% de las familias cumplieron con la escuela virtual")
    
    print("Se actualizo la celda de escuelas virtuales del AOR")


def editar_Hice(month,lugar):    

    if lugar=="colegio":
        sa=gspread.service_account(filename="C:/Users/gquintero/Desktop/python/pythontest-361720-6e048b13eddb.json")
    else:
        sa=gspread.service_account(filename="C:/Users/PC/Desktop/python/pythontest-361720-6e048b13eddb.json")
    
    google_sheet=sa.open("AOR Gustavo Quintero")
    hoja=google_sheet.worksheet("AÑO")
    col=hoja.find(month)
    hoja.update_cell(20,col.col,"ningun estudiante incumplio con actividades HICE este mes")
    
    print("Se actualizo la celda de HICE del AOR")


def editar_irrespeto(month,observadores,lugar):    
    comentario=""
    comentario_counter=""
    counter=0
    if len(observadores)==0:
        comentario_counter="Ningun estudiante tuvo anotacion en el observador por Irrespeto"
    else:
        for estudiante in observadores:
            for anotacion in observadores[estudiante]:
                if anotacion["codigo"]=="68.1.31" or anotacion["codigo"]=="68.1.36" or anotacion["codigo"]=="68.2.19" or anotacion["codigo"]=="68.2.21" or anotacion["codigo"]=="68.2.24":
                    counter+=1
                    comentario+=f"{estudiante},  "
            comentario_counter=str(counter)+"  "+comentario 
    
    if comentario_counter=="0  ":
        comentario_counter="Ningun estudiante tuvo anotacion en el observador por Irrespeto"

    if lugar=="colegio":
        sa=gspread.service_account(filename="C:/Users/gquintero/Desktop/python/pythontest-361720-6e048b13eddb.json")
    else:
        sa=gspread.service_account(filename="C:/Users/PC/Desktop/python/pythontest-361720-6e048b13eddb.json")
        
    google_sheet=sa.open("AOR Gustavo Quintero")
    hoja=google_sheet.worksheet("AÑO")
    col=hoja.find(month)
    hoja.update_cell(23,col.col,comentario_counter)
    print("Se actualizo la celda de irrespeto del AOR")


def editar_acosados(month,lugar):        

    if lugar=="colegio":
        sa=gspread.service_account(filename="C:/Users/gquintero/Desktop/python/pythontest-361720-6e048b13eddb.json")
    else:
        sa=gspread.service_account(filename="C:/Users/PC/Desktop/python/pythontest-361720-6e048b13eddb.json")
        
    google_sheet=sa.open("AOR Gustavo Quintero")
    hoja=google_sheet.worksheet("AÑO")
    col=hoja.find(month)
    hoja.update_cell(24,col.col,"No hay reportes de estudiantes victimas de acoso")
    print("Se actualizo la celda de acosados del AOR")


def editar_realizan_acoso(month,observadores,lugar):    
    comentario=""
    comentario_counter=""
    counter=0
    if len(observadores)==0:
        comentario_counter="No hay reportes de estudiantes que realicen acoso escolar"
    else:
        for estudiante in observadores:
            for anotacion in observadores[estudiante]:
                if anotacion["codigo"]=="68.2.23":
                    counter+=1
                    comentario+=f"{estudiante},  "
            comentario_counter=str(counter)+"  "+comentario 

    if comentario_counter=="0  ":
        comentario_counter="No hay reportes de estudiantes que realicen acoso escolar"

    if lugar=="colegio":
        sa=gspread.service_account(filename="C:/Users/gquintero/Desktop/python/pythontest-361720-6e048b13eddb.json")
    else:
        sa=gspread.service_account(filename="C:/Users/PC/Desktop/python/pythontest-361720-6e048b13eddb.json")
        
    google_sheet=sa.open("AOR Gustavo Quintero")
    hoja=google_sheet.worksheet("AÑO")
    col=hoja.find(month)
    hoja.update_cell(25,col.col,comentario_counter)
    print("Se actualizo la celda de realizar acoso del AOR")


def editar_invisibles(month,lugar):        

    if lugar=="colegio":
        sa=gspread.service_account(filename="C:/Users/gquintero/Desktop/python/pythontest-361720-6e048b13eddb.json")
    else:
        sa=gspread.service_account(filename="C:/Users/PC/Desktop/python/pythontest-361720-6e048b13eddb.json")
        
    google_sheet=sa.open("AOR Gustavo Quintero")
    hoja=google_sheet.worksheet("AÑO")
    col=hoja.find(month)
    hoja.update_cell(30,col.col,"No hay familias invisibles")
    print("Se actualizo la celda de Familias invisibles del AOR")


def editar_quejas(month,lugar):        

    if lugar=="colegio":
        sa=gspread.service_account(filename="C:/Users/gquintero/Desktop/python/pythontest-361720-6e048b13eddb.json")
    else:
        sa=gspread.service_account(filename="C:/Users/PC/Desktop/python/pythontest-361720-6e048b13eddb.json")
        
    google_sheet=sa.open("AOR Gustavo Quintero")
    hoja=google_sheet.worksheet("AÑO")
    col=hoja.find(month)
    hoja.update_cell(32,col.col,"No se recibieron quejas este mes")
    print("Se actualizo la celda de quejas del AOR")


def editar_capacitaciones(month,lugar):        

    if lugar=="colegio":
        sa=gspread.service_account(filename="C:/Users/gquintero/Desktop/python/pythontest-361720-6e048b13eddb.json")
    else:
        sa=gspread.service_account(filename="C:/Users/PC/Desktop/python/pythontest-361720-6e048b13eddb.json")
        
    google_sheet=sa.open("AOR Gustavo Quintero")
    hoja=google_sheet.worksheet("AÑO")
    col=hoja.find(month)
    hoja.update_cell(34,col.col,"Lenguaje de programacion Python")
    print("Se actualizo la celda de capacitaciones del AOR")
    

def open_inews2(lugar,ussername,password):
    if lugar=="colegio":
        PATH="C:/Users/gquintero/Desktop/python/chromedriver.exe"
    else:
        PATH="C:/Users/PC/Desktop/python/chromedriver.exe" 
    driver=webdriver.Chrome(PATH)   
    driver.get("https://inews2.interasesores.com.co/login/")
    search=driver.find_element(By.ID,value="user")
    search.send_keys(ussername)
    search=driver.find_element(By.ID,value="pass")
    search.send_keys(password)
    search.send_keys(Keys.RETURN)
    return driver


def close_inews2(driver,seconds):
    time.sleep(seconds)
    driver.close()


def get_name_escuela_funcionarios(driver,lugar):    
    if lugar=="colegio":
        coments_file_location="C:/Users/gquintero/Downloads"
    else:
        coments_file_location="C:/Users/PC/Downloads"
    driver.get("https://inews2.interasesores.com.co/cursos/")
    section=driver.find_element(By.ID, "main-content")
    contentcontainer=section.find_element(By.ID, "content-container")
    content=contentcontainer.find_element(By.ID, "content")
    article=content.find_element(By.TAG_NAME,"article")
    div=article.find_element(By.TAG_NAME,"div")
    escuela_actual=div.find_elements(By.TAG_NAME,"li")[0]
    a=escuela_actual.find_element(By.TAG_NAME,"a")
    nombre=a.find_element(By.TAG_NAME,"h3").text
    return nombre


def editar_escuela_funcionarios(month,lugar,nombre):        

    if lugar=="colegio":
        sa=gspread.service_account(filename="C:/Users/gquintero/Desktop/python/pythontest-361720-6e048b13eddb.json")
    else:
        sa=gspread.service_account(filename="C:/Users/PC/Desktop/python/pythontest-361720-6e048b13eddb.json")
    
    
    google_sheet=sa.open("AOR Gustavo Quintero")
    hoja=google_sheet.worksheet("AÑO")
    col=hoja.find(month)

    numero_escuela_anterior=hoja.cell(35,col.col-1).value[7:9]
    if nombre.startswith(f"MÓDULO {int(numero_escuela_anterior)+1}"):
        hoja.update_cell(35,col.col,nombre)
        print("Se actualizo la celda de nombre de escuela virtual del AOR")
    else:
        print("No ha salido una nueva escuela para funcionarios")
    


def main():
    mes=get_month("11") #"today" para el mes actual, "#" para el mes #.  Asi, "7" para el mes 7
    lugar="colegio"
    
    month=month_num_to_word(mes,2)
    observadores=get_observadores(mes,lugar)    
    
    editar_Hice(month,lugar)
    editar_uniforme(month,observadores,lugar)
    editar_tarde(month,observadores,lugar)
    editar_irrespeto(month,observadores,lugar)
    editar_acosados(month,lugar)
    editar_realizan_acoso(month,observadores,lugar)
    editar_escuela_virtual(month,lugar)
    editar_invisibles(month,lugar)
    editar_quejas(month,lugar)
    editar_capacitaciones(month,lugar)

    driver=open_inews2(lugar,config("inews2_usuario"),config("inews2_password"))
    nombre=get_name_escuela_funcionarios(driver,lugar)
    close_inews2(driver,3)
    editar_escuela_funcionarios(month,lugar,nombre)

if __name__ == "__main__":
    main()