import os
import shutil
import subprocess
import tempfile
import tkinter as tk
from tkinter import messagebox

# Configuraciones
shared_folder = r"Z:\InSoft\Semaforo\sicore"
local_folder = r"C:\Program Files (x86)\S.I.Ap\AFIP\sicore"
lock_file = os.path.join(shared_folder, "lock.txt")
mdb_filename = "SICORE.mdb"
db3_filename = "Sicore.db3"
siap_exe = r"C:\Program Files (x86)\S.I.Ap\AFIP\siap.exe"

def is_locked():
    return os.path.exists(lock_file)

def create_lock():
    with open(lock_file, 'w') as f:
        f.write(os.environ['COMPUTERNAME'])

def remove_lock():
    if os.path.exists(lock_file):
        os.remove(lock_file)

def get_lock_owner():
    if is_locked():
        with open(lock_file, 'r') as f:
            return f.read().strip()
    return None

def copy_from_shared():
    shutil.copy2(os.path.join(shared_folder, mdb_filename), local_folder)
    shutil.copy2(os.path.join(shared_folder, db3_filename), local_folder)

def copy_to_shared():
    shutil.copy2(os.path.join(local_folder, mdb_filename), shared_folder)
    shutil.copy2(os.path.join(local_folder, db3_filename), shared_folder)


def run_siap():
    # Creamos un script temporal de PowerShell que ejecute siap.exe como admin y guarde el PID
    ps_script = """
    $p = Start-Process -FilePath '{}'-PassThru -Verb RunAs
    $p.Id | Out-File '{}'
    $p.WaitForExit()
    """.format(siap_exe.replace('"', '""'), os.path.join(tempfile.gettempdir(), "siap_pid.txt"))

    script_path = os.path.join(tempfile.gettempdir(), "run_siap_admin.ps1")
    with open(script_path, 'w') as f:
        f.write(ps_script)

    # Ejecutamos el script sin mostrar ventana
    return subprocess.Popen(
        ["powershell", "-WindowStyle", "Hidden", "-ExecutionPolicy", "Bypass", "-File", script_path],
        creationflags=subprocess.CREATE_NO_WINDOW
    )

def wait_for_siap_exit(proc):
    proc.wait()

def iniciar_proceso():
    if is_locked():
        owner = get_lock_owner()
        messagebox.showwarning("Acceso denegado", f"SICORE está siendo usado por: {owner}")
        return

    try:
        create_lock()
        status_label.config(text="📥 Copiando base de datos desde red...", fg="blue")
        root.update()

        copy_from_shared()

        status_label.config(text="🚀 Ejecutando SIAP...", fg="green")
        root.update()

        siap_process = run_siap()
        wait_for_siap_exit(siap_process)

        status_label.config(text="📤 Copiando base de datos a la red...", fg="blue")
        root.update()

        copy_to_shared()

    except Exception as e:
        messagebox.showerror("Error", str(e))
    finally:
        remove_lock()
        status_label.config(text="✅ SICORE libre", fg="green")

def actualizar_estado():
    if is_locked():
        owner = get_lock_owner()
        status_label.config(text=f"🔴 SICORE en uso por: {owner}", fg="red")
        start_button.config(state="disabled")
    else:
        status_label.config(text="🟢 SICORE libre", fg="green")
        start_button.config(state="normal")
    root.after(3000, actualizar_estado)  # Actualiza cada 3 segundos

# Interfaz gráfica
root = tk.Tk()
root.title("InSoft - Acceso exclusivo a SICORE")
root.geometry("400x200")
root.iconbitmap("insoft.ico")

status_label = tk.Label(root, text="", font=("Arial", 14))
status_label.pack(pady=30)

start_button = tk.Button(root, text="Usar SICORE", font=("Arial", 12), command=iniciar_proceso)
start_button.pack(pady=10)

actualizar_estado()

root.mainloop()
