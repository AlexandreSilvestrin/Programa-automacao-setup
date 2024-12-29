import tkinter as tk
from tkinter import scrolledtext, messagebox


def get_local_version(file_path=r'C:\ProgramData\xy-auto\version.txt'):
    try:
        with open(file_path, 'r') as file:
            version_str = file.read().strip()
            return version.parse(version_str)
    except FileNotFoundError:
        print(f"Arquivo {file_path} não encontrado. Definindo versão como '0.0'")
        return version.parse('0.0')

def preserve_database(db_path=r'C:\ProgramData\xy-auto\exe.win-amd64-3.12\resources\data\BANCOCNPJ.xlsx', backup_path=r'BANCOCNPJ_backup.xlsx'):
    if os.path.exists(db_path):
        shutil.copy2(db_path, backup_path)
        print(f"Banco de dados preservado em {backup_path}")

def restore_database(db_path=r'C:\ProgramData\xy-auto\exe.win-amd64-3.12\resources\data\BANCOCNPJ.xlsx', backup_path=r'BANCOCNPJ_backup.xlsx'):
    if os.path.exists(backup_path):
        shutil.copy2(backup_path, db_path)
        print(f"Banco de dados restaurado de {backup_path}")
        os.remove(backup_path)

def update_local_version(new_version_str, file_path=r'C:\ProgramData\xy-auto\version.txt'):
    with open(file_path, 'w') as file:
        file.write(new_version_str)
    print(f"Versão local atualizada para {new_version_str}.")

def get_latest_release_info(repo_owner, repo_name):
    tag_name, download_url = None, None
    try:
        url = f"https://api.github.com/repos/{repo_owner}/{repo_name}/releases/latest"
        response = requests.get(url)
        response.raise_for_status()
        latest_release = response.json()
        tag_name = latest_release['tag_name']
        download_url = latest_release['assets'][0]['browser_download_url']
        return tag_name, download_url , True
    except:
        return tag_name, download_url , False


def download_new_version(download_url, version_str, download_dir=r'C:\ProgramData\xy-auto\downloads'):
    if not os.path.exists(download_dir):
        os.makedirs(download_dir)
    local_filename = os.path.join(download_dir, f'app_{version_str}.rar')
    
    with requests.get(download_url, stream=True) as r:
        r.raise_for_status()
        total_size = int(r.headers.get('content-length', 0))
        block_size = 8192
        with open(local_filename, 'wb') as f:
            for chunk in r.iter_content(chunk_size=block_size):
                f.write(chunk)
    print(f"Download completo: {local_filename}")
    return local_filename

def extract_rar_file(rar_file_path, extract_to_dir=r'C:\ProgramData\xy-auto'):
    # Lista de caminhos possíveis onde o WinRAR pode estar instalado
    possible_paths = [
        r"C:\Program Files\WinRAR\WinRAR.exe",
        r"C:\Program Files (x86)\WinRAR\WinRAR.exe",
        r"D:\Program Files\WinRAR\WinRAR.exe",
        r"D:\Program Files (x86)\WinRAR\WinRAR.exe"
    ]
    
    winrar_path = None
    for path in possible_paths:
        if os.path.exists(path):
            winrar_path = path
            break

    if not winrar_path:
        raise FileNotFoundError("WinRAR não encontrado em nenhum dos caminhos padrão.")

    
    if not os.path.exists(extract_to_dir):
        os.makedirs(extract_to_dir)
    
    # Usando o WinRAR no caminho encontrado para extrair o arquivo
    subprocess.run([winrar_path, 'x', '-y', rar_file_path, extract_to_dir], check=True)
    
    print(f"Extração completa em: {extract_to_dir}")
    


def execute_program():
    program_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    program_path = os.path.join(program_dir, r'C:\ProgramData\xy-auto', 'exe.win-amd64-3.12', 'app.exe')
    os.chdir(os.path.dirname(program_path))
    subprocess.run([program_path])
    print(f"Executando {program_path}")

def remove_downloads_folder(download_dir=r'C:\ProgramData\xy-auto\downloads'):
    try:
        if os.path.exists(download_dir):
            shutil.rmtree(download_dir)
            print(f"Pasta '{download_dir}' excluída com sucesso.")
        else:
            print(f"Pasta '{download_dir}' não encontrada.")
    except Exception as e:
        print(f"Erro ao tentar excluir a pasta '{download_dir}': {e}")


def check_for_update():
    repo_owner = 'AlexandreSilvestrin'
    repo_name = 'Programa-automacao'

    current_version = get_local_version()
    print(f"Versão atual: {current_version}")

    try:
        latest_version_str, download_url, cond = get_latest_release_info(repo_owner, repo_name)
        if cond:
            latest_version = version.parse(latest_version_str)
            print(f"Última versão disponível: {latest_version}")
        else:
            return None, None, False
    except requests.HTTPError as e:
        print(f"Erro ao obter informações da última versão: {e}")
        return None, None, 'NET'

    if latest_version > current_version:
        print(f"Nova versão {latest_version} disponível.")
        return latest_version_str, download_url, current_version
    else:
        print("Você já possui a versão mais recente.")
        return None, None, current_version

def center_window(window, width=500, height=400):
    
    # Obtém a largura e altura da tela
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()

    # Calcula a posição da janela para centralizá-la
    position_x = int((screen_width / 2) - (width / 2))
    position_y = int((screen_height / 2) - (height / 2))

    # Define a geometria da janela com a posição centralizada
    window.geometry(f'{width}x{height}+{position_x}+{position_y}')


def main():
    
    # Cria a janela principal
    window = tk.Tk()
    window.title("Executar e Atualizar")
    window.geometry("500x400")
    window.resizable(False, False)
    center_window(window)

    status_label = tk.Label(window, text="Bem-vindo", font=("Arial", 14, "bold"))
    status_label.pack(pady=20)

    log_box = scrolledtext.ScrolledText(window, width=60, height=15)
    log_box.pack(pady=20)

    def carregarDP():
        global requests, os, subprocess, shutil, version, sys
        import requests
        import os
        import subprocess
        import shutil
        from packaging import version
        import sys
        clear_log()
        check_update_button.pack(side=tk.RIGHT, padx=10, pady=10)
        open_program_button.pack(side=tk.RIGHT, padx=10, pady=10)

    def log_message(message):
        log_box.insert(tk.END, message + "\n")
        log_box.see(tk.END)
        window.update_idletasks()

    def clear_log():
        log_box.delete(1.0, tk.END)
    
    def delete_atual():
        program_dir = 'C:/ProgramData/xy-auto/exe.win-amd64-3.12'  # Caminho da pasta a ser deletada
        try:
            if os.path.exists(program_dir):
                shutil.rmtree(program_dir)
        except Exception as e:
            log_message(f"Erro ao tentar excluir a pasta '{program_dir}': {e}")

    def start_update():
        clear_log()
        latest_version_str, download_url, current_version = check_for_update()
        if latest_version_str and download_url:
            response = messagebox.askyesno("Confirmação", f"Nova atualização disponível:\nVersão atual: v{current_version}\nNova versão: v{latest_version_str}\nDeseja atualizar?")
        
            if response:
                try:
                    log_message(f"Atualizando para a versão {latest_version_str}...")
                    downloaded_file = download_new_version(download_url, latest_version_str)
                    preserve_database()
                    delete_atual()
                    extract_rar_file(downloaded_file)
                    restore_database()
                    update_local_version(latest_version_str)
                    remove_downloads_folder()
                    log_message("Atualização concluída.")
                    messagebox.showinfo("Finalizado", f"Atualização para versão v{latest_version_str} concluída.")
                except Exception as e:
                    log_message(f"Erro durante a atualização: {e}")
            else:
                log_message(f"Uma nova versão disponivel: {latest_version_str}")
        else:
            messagebox.showinfo("Resposta", f"Seu programa ja esta na ultima atualização v{current_version}")
            log_message(f"Nenhuma atualização disponível.\nVersão atual: v{current_version}.")

    def open_program():
        window.destroy()
        execute_program()

    check_update_button = tk.Button(window, text="Verificar Atualização", command=start_update)
    open_program_button = tk.Button(window, text="Abrir Programa", command=open_program)

    check_update_button.pack_forget()
    open_program_button.pack_forget()

    log_message("Aguarde carregando...")
    window.after(2000, carregarDP)
    window.mainloop()


if __name__ == "__main__":
    main()