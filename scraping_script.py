import json
import shutil
import time
import re
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openai

# Configurar la clave de API
openai.api_key = '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'
# Realizar una solicitud a la API de ChatGPT
response = openai.Completion.create(
  engine='gpt-3.5-turbo',
  prompt='Pregunta: ¿Por qué es beneficioso hacer un bootcamp en CORE Code School?',
  max_tokens=50
)

# Obtener la respuesta generada
answer = response.choices[0].text.strip()

# Imprimir la respuesta
print(answer)
# Ruta al archivo user_details.json
archivo_user_details = "user_details.json"
# Abrir el archivo user_details.json para lectura
with open(archivo_user_details, "r") as archivo:
        # Cargar los datos del archivo JSON
        detalles_usuario = json.load(archivo)
        
        # Recuperar los valores de usuario y contraseña
        usuario_json = detalles_usuario["LOGIN_USERNAME"]
        contrasena = detalles_usuario["LOGIN_PASSWORD"]
# Solicitar al usuario el número de sprint
while True:
    sprint = input("Ingrese el número de sprint: ")
    # Verificar si el sprint es un número entero
    if sprint.isdigit():
        sprint = int(sprint)
        break
    else:
        print("Error: El sprint debe ser un número entero.")

# Solicitar al usuario su nombre de usuario en GitHub
usuario = input("Ingrese su nombre de usuario en GitHub: ")

# Construir la URL de acuerdo al sprint y usuario ingresados por el usuario
url_base = "https://github.com/kesslertopaz/STI-v3/issues?q=is%3Aopen+is%3Aissue"
parte_sprint = f"+label%3A%22Sprint+{sprint}%22"
parte_usuario = f"+assignee%3A{usuario}"
url = url_base + parte_sprint + parte_usuario

try:
    # Configuración de Selenium
    chrome_options = Options()
    #chrome_options.add_argument("--headless")  # Para ejecutar Chrome en modo Headless
    chrome_service = ChromeService("C:\\Users\\josem\\OneDrive\\Documentos\\GitHub\\QA-Design-Automation\\chromedriver-win64\\chromedriver.exe")

    # Inicializar el navegador
    driver = webdriver.Chrome(service=chrome_service, options=chrome_options)
    
    # Navegar a la URL generada
    driver.get(url)

    # Completar el formulario de inicio de sesión
    username_field = driver.find_element(By.NAME, "login")  
    username_field.send_keys(usuario)

    password_field = driver.find_element(By.NAME, "password")  
    password_field.send_keys(contrasena)
    password_field.send_keys(Keys.RETURN)

    # Esperar hasta que aparezca el campo para el segundo factor
    wait = WebDriverWait(driver, 10)
    two_factor_field = wait.until(EC.presence_of_element_located((By.NAME, "app_otp")))

    # Guardar las ventanas actuales
    main_window = driver.current_window_handle
    all_windows = driver.window_handles

    # Cambiar el enfoque a la ventana del segundo factor
    for window in all_windows:
        if window != main_window:
            driver.switch_to.window(window)

    # Esperar un tiempo adicional si es necesario para las acciones en la ventana del segundo factor

    # Cambiar el enfoque nuevamente a la ventana principal
    driver.switch_to.window(main_window)

    # Esperar 30 segundos
    time.sleep(30)

    # Cambiar el enfoque nuevamente a la ventana principal
    driver.switch_to.window(main_window)
    
    # Verificar que la URL actual sea la esperada después de los 30 segundos
    current_url = driver.current_url
    all_cases = {}

    if current_url == current_url:
        print("Estás en la URL correcta después de 30 segundos.")
        # Obtener información del caso
        case_elements = driver.find_elements(By.CSS_SELECTOR, 'a[id^="issue_"]')

        for case_element in case_elements:
            case_number_match = re.search(r'issue_(\d+)_link', case_element.get_attribute("id"))
            if case_number_match:
                case_number = case_number_match.group(1)
                case_name = case_element.text.replace('"', '').replace('/', ' ')
                sprint_match = re.search(r'Sprint\+(\d+)', current_url)
                sprint_number = sprint_match.group(1) if sprint_match else None
                user_match = re.search(r'assignee%3A(\w+)', current_url)
                username = user_match.group(1) if user_match else None
                case_url = case_element.get_attribute("href")

                case_info = {
                    "Numero_de_Caso": case_number,
                    "Nombre_del_Caso": case_name,
                    "Sprint": sprint_number,
                    "Usuario": username,
                    "URL_del_Caso": case_url
                }

                all_cases[case_number] = case_info

                template_path = "Template/Release.12.Template.xlsx"
                new_filename = f"Release.{sprint_number}.{case_name} - {case_number}.xlsx"
                destination_path = f"Template/{new_filename}"
                shutil.copy(template_path, destination_path)
                print(f"Se ha creado la copia para el caso {case_number}: {new_filename}")
    else:
        print("La URL no coincide con la esperada.")
        
    total_cases_processed = len(all_cases)
    print(f"\nNúmero total de casos procesados: {total_cases_processed}")
    print(json.dumps(list(all_cases.values()), indent=2))
    
except Exception as e:
    print(f"Error durante la ejecución: {str(e)}")

finally:
    # Cerrar el navegador al finalizar
    driver.quit()
