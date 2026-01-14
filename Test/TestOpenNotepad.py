# Teste utilizando pywinauto

from pywinauto import Application
from pywinauto.keyboard import send_keys
import time

# Abre o Bloco de Notas
app = Application(backend="uia").start("notepad.exe")

# Aguarda a janela existir
time.sleep(1)

# Conecta à janela do Notepad
janela = app.window(class_name="Notepad")

# Garante foco lógico (não visual)
janela.wait("visible", timeout=10)
janela.set_focus()

send_keys("Agora sim funciona corretamente!{ENTER}"
          "Este texto foi digitado via pywinauto.{ENTER}",
          with_spaces=True
)

