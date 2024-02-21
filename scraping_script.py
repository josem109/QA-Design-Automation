from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import shutil
import time
import json
import re

# Configuración de Selenium
chrome_options = Options()
#chrome_options.add_argument("--headless")  # Para ejecutar Chrome en modo sin cabeza
chrome_service = ChromeService("C:\\Users\\josem\\OneDrive\\Documentos\\GitHub\\QA-Design-Automation\\chromedriver-win64\\chromedriver.exe")
# Credenciales
usuario = "jpinto@ktmc.com"
contrasena = "*Gr4ndl0rd*"


try:
    # Inicializar el navegador
    driver = webdriver.Chrome(service=chrome_service, options=chrome_options)

    # Navegar a la URL que requiere autenticación
    url = "https://github.com/kesslertopaz/STI-v3/issues?q=is%3Aopen+is%3Aissue+label%3A%22Sprint+37%22+assignee%3Ajpintoktmc+-label%3A%22QA+Done%22"
    driver.get(url)

    username_field = driver.find_element(By.NAME, "login")  # Reemplaza con el nombre real del campo
    username_field.send_keys(usuario)

    # Encontrar y completar el campo de contraseña
    password_field = driver.find_element(By.NAME, "password")  # Reemplaza con el nombre real del campo
    password_field.send_keys(contrasena)

    # Enviar el formulario de inicio de sesión
    password_field.send_keys(Keys.RETURN)

    # Esperar hasta que aparezca el campo para el segundo factor (puedes ajustar el tiempo según sea necesario)
    wait = WebDriverWait(driver, 10)
    two_factor_field = wait.until(EC.presence_of_element_located((By.NAME, "app_otp")))

    # Guardar las ventanas actuales
    main_window = driver.current_window_handle
    all_windows = driver.window_handles

    # Realizar acciones adicionales en la ventana del segundo factor (puedes agregar aquí cualquier código necesario)
    # ...

    # Cambiar el enfoque a la ventana del segundo factor
    for window in all_windows:
        if window != main_window:
            driver.switch_to.window(window)

    # Esperar un tiempo adicional si es necesario para las acciones en la ventana del segundo factor
    # driver.implicitly_wait(5)

    # Realizar más acciones en la ventana del segundo factor si es necesario
    # ...

    # Cambiar el enfoque nuevamente a la ventana principal
    driver.switch_to.window(main_window)

      # Esperar 30 segundos (puedes ajustar el tiempo según sea necesario)
    time.sleep(30)

    # Cambiar el enfoque nuevamente a la ventana principal
    driver.switch_to.window(main_window)
    # Obtener el título de la página
 # Verificar que la URL actual sea la esperada después de los 30 segundos
    current_url = driver.current_url
    expected_url = "https://github.com/kesslertopaz/STI-v3/issues?q=is%3Aopen+is%3Aissue+label%3A%22Sprint+37%22+assignee%3Ajpintoktmc+-label%3A%22QA+Done%22"
    # Diccionario para almacenar los objetos JSON
    all_cases = {}
    if current_url == expected_url:
        print("Estás en la URL correcta después de 30 segundos.")
         # Obtener información del caso
    case_elements = driver.find_elements(By.CSS_SELECTOR, 'a[id^="issue_"]')

    for case_element in case_elements:
          case_number_match = re.search(r'issue_(\d+)_link', case_element.get_attribute("id"))
          if case_number_match:
            case_number = case_number_match.group(1)
            case_name = case_element.text.replace('"', '').replace('/', ' ')
            # Obtener el número de sprint de la URL
            sprint_match = re.search(r'Sprint\+(\d+)', current_url)
            sprint_number = sprint_match.group(1) if sprint_match else None

            # Obtener el nombre de usuario de la URL
            user_match = re.search(r'assignee%3A(\w+)', current_url)
            username = user_match.group(1) if user_match else None
            
            case_url = case_element.get_attribute("href")

            # Crear un diccionario con la información del caso
            case_info = {
                "Numero_de_Caso": case_number,
                "Nombre_del_Caso": case_name,
                "Sprint": sprint_number,
                "Usuario": username,
                "URL_del_Caso": case_url
            }

            # Agregar el objeto JSON al diccionario usando el número de caso como clave
            all_cases[case_number] = case_info

            # Imprimir el objeto JSON
            import json
            #print(json.dumps(case_info, indent=2))
            # Crear una copia del archivo Release.12.Template.xlsx
            template_path = "Template/Release.12.Template.xlsx"
            new_filename = f"Release.{sprint_number}.{case_name} - {case_number}.xlsx"
            destination_path = f"Template/{new_filename}"  # Ajusta la carpeta de destino según tus necesidades

            shutil.copy(template_path, destination_path)
            print(f"Se ha creado la copia para el caso {case_number}: {new_filename}")
    else:
        print("La URL no coincide con la esperada.")
    # Imprimir la cantidad de casos procesados
    total_cases_processed = len(all_cases)
    print(f"\nNúmero total de casos procesados: {total_cases_processed}")
    # Imprimir el diccionario de objetos JSON al final
    print(json.dumps(list(all_cases.values()), indent=2))
except Exception as e:
    print(f"Error durante la ejecución: {str(e)}")

finally:
    # Cerrar el navegador al finalizar
    driver.quit()