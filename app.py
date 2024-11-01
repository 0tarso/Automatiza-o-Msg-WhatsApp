import openpyxl
from urllib.parse import quote
import webbrowser
import time
import pyautogui

workbook = openpyxl.load_workbook("clientes.xlsx")
customer_sheet = workbook["Plan1"]


# def open_browser():
#     webbrowser.open("https://web.whatsapp.com")
#     time.sleep(15)


def send_message():

    # para cada linha na tabela
    for row in customer_sheet.iter_rows(min_row=2):

        # o index representa a coluna da tabela
        name = row[0].value
        phone = row[1].value

        message = f"Teste {name}, seu numero é {phone}"

        message_encode = quote(message)

        print(message_encode)

        try:
            url_wpp = (
                f"https://web.whatsapp.com/send?phone={phone}&text={message_encode}"
            )

            webbrowser.open(url_wpp)
            time.sleep(10)

            send_icon = pyautogui.locateCenterOnScreen("arrowIcon.png")
            time.sleep(3)

            pyautogui.click(send_icon[0], send_icon[1])
            time.sleep(5)

            pyautogui.hotkey("ctrl", "w")
            time.sleep(5)

        except:
            print(f"Erro ao enviar para {name}")

            with open("erros.csv", "a", newline="", encoding="utf-8") as arq:
                arq.write(f"{name}, {phone}")

    print("Envios Concluidos")


send_message()
