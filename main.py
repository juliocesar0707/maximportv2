# main.py
import sys
import os

print("Iniciando Max Import 2.0...")
print("Carregando Interface Gráfica Moderna...")

# Apenas importa e executa o app.py, que agora é o módulo principal do sistema
import app

if __name__ == "__main__":
    app_instance = app.MaxImportApp()
    app_instance.mainloop()