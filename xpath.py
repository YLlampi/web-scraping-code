from time import sleep, time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
import pandas as pd

inicio = time()

def get_avance_financiero(a, b):
    a = int(a.replace('S/ ', '').replace(',', ''))
    b = int(b.split('.')[0].replace(',', ''))
    c = round(float(a/b), 3)*100
    return c


def get_codigos():
    lista_codigo_unico = []
    f1 = open("./codigos.txt")

    while True:
        line = f1.readline()
        if ("" == line):
            break;
        lista_codigo_unico.append(line.replace("\n", ""))

    return lista_codigo_unico


opts = Options()
#opts.add_argument("Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36")
opts.add_argument('--start-maximized')
opts.add_argument('--disable-extensions')
driver = webdriver.Chrome('./chromedriver_linux64/chromedriver', chrome_options=opts)

url_website = "https://ofi5.mef.gob.pe/ssi/ssi/Index"
driver.get(url_website)


lista_codigo_unico = get_codigos()
# Primera Capa
lista_nombre_inversion = []
lista_costo_actualizado_inversion_a = []
lista_costo_total_inversion_d = []
lista_tiene_registro = []
lista_registro_ET = []
lista_registrado_PMI = []

# Segunda Capa
lista_pim = []
lista_devengado_acumulado = []
# Formula Matematica
lista_avance_financiero = []

# Tercera Capa N1
lista_fecha_ultima_situacion_inversion = []

# Tercera Capa N2
lista_fecha_ultima_actualizacion_declaracion_inversion = []
lista_avance_ejecucion_inversion = []
lista_avance_fisico = []

# Cuarta Capa
lista_ultima_valorizacion = []
lista_real_acumulado_porcentaje = []
lista_real_acumulado_soles = []


print("[OBTENIENDO DATOS...]")

for codigo in lista_codigo_unico:
    driver.get(url_website)

    input_codigo = driver.find_element('xpath', '//input[@id="txt_cu"]')
    input_codigo.send_keys(codigo)

    search_button_click = driver.find_element('xpath', '//span[@class="pt-1 align-items-start btn_bus"]')
    search_button_click.click()

    link_tag = ""

    band = True

    try:
        link_green_button = WebDriverWait(driver, 4).until(EC.element_to_be_clickable((By.XPATH, '//td[@id="td_indseg"]/a')))
        link_tag = link_green_button.get_attribute("href")

    except Exception as e:
        band = False
        lista_pim.append(" ")
        lista_devengado_acumulado.append(" ")
        lista_avance_financiero.append(" ")
        lista_fecha_ultima_actualizacion_declaracion_inversion.append(" ")
        lista_avance_ejecucion_inversion.append(" ")
        lista_avance_fisico.append(" ")
        lista_fecha_ultima_situacion_inversion.append(" ")
        lista_ultima_valorizacion.append(" ")
        lista_real_acumulado_porcentaje.append(" ")
        lista_real_acumulado_soles.append(" ")
        print(e)
    finally:
        # Instrucciones xpath
        nombre = driver.find_element('xpath', '//td[@id="td_nominv"]').text
        lista_nombre_inversion.append(nombre)

        costo_actualizado = driver.find_element('xpath', '//td[@id="val_cta"]').text
        lista_costo_actualizado_inversion_a.append(costo_actualizado)

        costo_total = driver.find_element('xpath', '//td[@id="td_mtototal"]').text
        lista_costo_total_inversion_d.append(costo_total)
        
        # -------------------
        is_registro = driver.find_element('xpath', '//td[@id="td_indseg"]').text
        lista_tiene_registro.append(is_registro.strip())
        # -------------------

        registro_ET = driver.find_element('xpath', '//td[@id="td_indet"]').text
        lista_registro_ET.append(registro_ET.strip())

        registrado_PMI = driver.find_element('xpath', '//td[@id="td_indpmi"]').text
        lista_registrado_PMI.append(registrado_PMI.strip())

    if band == True:
        driver.get(link_tag)

        link_tag_detalle = ""
        band_2 = True
        try:
            boton_detalle_avance = WebDriverWait(driver, 6).until(
                EC.element_to_be_clickable((By.XPATH, '//span[@id="det_avan"]/a')),
            )
            link_tag_detalle = boton_detalle_avance.get_attribute("href")
        except Exception as e:
            band_2 = False
            lista_fecha_ultima_actualizacion_declaracion_inversion.append("-")
            lista_avance_ejecucion_inversion.append("0")
            lista_avance_fisico.append("0")
            print(e)


        # Link Ejecucion Fisica
        boton_ejecucion_fisica = driver.find_element('xpath', '//a[@id="detef"]')
        link_tag_ejecucion_fisica = boton_ejecucion_fisica.get_attribute("href")


        pim = driver.find_element('xpath', '//td[@id="val_pim"]').text
        lista_pim.append(pim.replace('S/ ', ''))        

        devengado_fake = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//a[@id="transp"]//img')),
        )
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, '//td[@id="val_efin"]')),
        )

        devengado_acumulado = driver.find_element('xpath', '//td[@id="val_efin"]').text
        avance_financiero = get_avance_financiero(devengado_acumulado, costo_actualizado)
        lista_avance_financiero.append(avance_financiero)

        lista_devengado_acumulado.append(devengado_acumulado.replace('S/ ', ''))


        # Tercera Capa N1
        boton_fecha_ultima_declaracion = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//td[@id="situ_nvo"]//i')),
        )
        boton_fecha_ultima_declaracion.click()

        boton_click_cerrar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//div[@class="modal-footer"]//button[2]')),
        )

        try:
            fecha_ultima_declaracion = driver.find_element('xpath', '//tbody[@id="tb_situ"]//td').text
        except Exception as e:
            fecha_ultima_declaracion = "-"
            print(e)
        finally:
            lista_fecha_ultima_situacion_inversion.append(fecha_ultima_declaracion)
            boton_click_cerrar.click()

        if band_2 == True:
            driver.get(link_tag_detalle)

            try:
                WebDriverWait(driver, 30).until(
                    EC.presence_of_all_elements_located((By.XPATH, '//tbody[@id="tbody_01a"]/tr')),
                )


                fecha_declaracion = driver.find_element('xpath', '//td[@id="fec_decla"]').text
                lista_fecha_ultima_actualizacion_declaracion_inversion.append(fecha_declaracion)

                avance_ejecucion_inversion = driver.find_element('xpath', '//td[@id="td_cab07"]').text
                lista_avance_ejecucion_inversion.append(avance_ejecucion_inversion.replace(' %', ''))
                
                avance_fisico = driver.find_element('xpath', '//td[@id="por_afis"]').text
                lista_avance_fisico.append(avance_fisico.replace(' %', ''))
            except Exception as e:
                print(e)


        # Cuarta capa
        driver.get(link_tag_ejecucion_fisica)
        
        band_3 = 0
        
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, '//tbody[@id="tbody_01"]/tr')),
            )
            band_3 = 1
            lista_ultima_valorizacion.append("-")
            lista_real_acumulado_porcentaje.append("-")
            lista_real_acumulado_soles.append("-")
        except Exception as e:
            #print(e)

        #if(band_3 == 0):

            WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, '//tbody[@id="tbody_01a"]/tr')),
            )


            boton_ver = driver.find_element('xpath', '//tbody[@id="tbody_01a"]//a')
            link_tag_boton_ver = boton_ver.get_attribute("href")
            
            driver.get(link_tag_boton_ver)

            WebDriverWait(driver, 30).until(
                EC.presence_of_all_elements_located((By.XPATH, '//tbody[@id="tbody_01a"]/tr')),
            )
            
            boton_avance_fisico = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, '//a[@id="tab_avfi"]')),
            )

            boton_avance_fisico.click()

            try:
                botonClick_Fake = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//tr[@id="0"]/td[5]/a[2]/img')),
                )
                """
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//tr[@id="0"]/td[1]')),
                )
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//tr[@id="0"]/td[4]')),
                )
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//tr[@id="0"]/td[5]')),
                )
                """
                ultima_valorizacion = driver.find_element('xpath', '//tbody[@id="tbody_01b"]/tr/td[1]').text
                lista_ultima_valorizacion.append(ultima_valorizacion.replace("\n",""))

                real_acumulado_porcentaje = driver.find_element('xpath', '//tbody[@id="tbody_01b"]/tr/td[4]').text
                lista_real_acumulado_porcentaje.append(real_acumulado_porcentaje.replace("%", "").replace("\n",""))

                real_acumulado_soles = driver.find_element('xpath', '//tbody[@id="tbody_01b"]/tr/td[5]').text
                lista_real_acumulado_soles.append(real_acumulado_soles.replace("\n","").strip())
            except Exception as e:
                lista_ultima_valorizacion.append("-")
                lista_real_acumulado_porcentaje.append("-")
                lista_real_acumulado_soles.append("-")
                print(e)
                    
    
print("[DATOS OBTENIDOS CORRECTAMENTE]")
print("[GENERANDO ARCHIVO...]")

data = {
    'CODIGO UNICO': lista_codigo_unico,
    'NOMBRE PIP': lista_nombre_inversion,
    'COSTO ACTUALIZADO DE LA INVERSION (A)': lista_costo_actualizado_inversion_a,
    'COSTO TOTAL DE INVERSION (D=A+B+C)': lista_costo_total_inversion_d,
    'PIM 2022': lista_pim,
    'DEVENGADO ACUMULADO AL 2022': lista_devengado_acumulado,
    'AVANCE FINANCIERO (E/D)': lista_avance_financiero,
    'TIENE REGISTRO DEL FORMATO 12-B': lista_tiene_registro,
    'FECHA ULTIMA ACTUALIZACION DE DECLARACION DE INVERSION': lista_fecha_ultima_actualizacion_declaracion_inversion,
    'AVANCE EJECUCION INVERSION': lista_avance_ejecucion_inversion,
    'AVANCE FISICO': lista_avance_fisico,
    'FECHA DE ULTIMA SITUACION DE INVERSION': lista_fecha_ultima_situacion_inversion,
    'REGISTRO ET': lista_registro_ET,
    'REGISTRADO EN EL PMI': lista_registrado_PMI,
    'ULTIMA VALORIZACION': lista_ultima_valorizacion,
    'REAL ACUMULADO (%)': lista_real_acumulado_porcentaje,
    'REAL ACUMULADO (S/.)': lista_real_acumulado_soles,
}

try:
    df_exportar = pd.DataFrame(data)
    df_exportar.to_excel('OPMI_OMACHA.xlsx', index=False)
    print("[ARCHIVO GENERADO]")
except Exception as e:
    print(e)
    print("[ERROR Al GENERAR ARCHIVO]")


driver.close()
fin = time()
print(f'Tiempo: {round((fin-inicio)/60, 1)}, minutos.')
print("[FIN PROGRAMA]")
