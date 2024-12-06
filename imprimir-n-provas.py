import os  # Importa a biblioteca para manipular arquivos e pastas
import platform  # Detecta o sistema operacional
import time  # Fornece funções para manipular tempo, como pausas
import subprocess  # Permite executar comandos do sistema operacional
from tkinter import Tk, filedialog, messagebox, StringVar, Label, Button, OptionMenu, Canvas, PhotoImage, ttk

# Função para listar as impressoras disponíveis no sistema
def listar_impressoras():
    """Lista as impressoras instaladas no sistema."""
    impressoras = []
    if os_name == "Windows":  # Verifica se o sistema operacional é Windows
        import win32print  # Biblioteca para manipular impressoras no Windows
        impressoras = [imp[2] for imp in win32print.EnumPrinters(2)]  # Obtém os nomes das impressoras
    elif os_name == "Linux":  # Verifica se o sistema operacional é Linux
        try:
            result = subprocess.run(["lpstat", "-a"], stdout=subprocess.PIPE, text=True)  # Executa o comando lpstat
            impressoras = [line.split()[0] for line in result.stdout.strip().split("\n")]  # Processa o output
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao listar impressoras no Linux: {e}")  # Mostra mensagem de erro
    return impressoras  # Retorna a lista de impressoras

# Função para imprimir arquivos no Linux
def imprimir_linux(caminho_arquivo, impressora):
    """Imprime um arquivo no Linux em uma impressora específica."""
    try:
        subprocess.run(["lp", "-d", impressora, caminho_arquivo], check=True)  # Usa o comando lp para impressão
    except Exception as e:
        raise RuntimeError(f"Erro ao imprimir {caminho_arquivo} na impressora {impressora}: {e}")  # Exceção em caso de erro

# Função para imprimir arquivos no Windows
def imprimir_windows(caminho_arquivo, impressora):
    """Imprime um arquivo no Windows em uma impressora específica."""
    import win32print  # Biblioteca para manipular impressoras no Windows
    win32print.SetDefaultPrinter(impressora)  # Define a impressora padrão
    import win32api  # Biblioteca para executar comandos no Windows
    win32api.ShellExecute(0, "print", caminho_arquivo, None, ".", 0)  # Executa o comando de impressão

# Função para gerenciar a impressão de todos os arquivos na pasta escolhida
def imprimir_arquivos():
    if not caminho:  # Verifica se uma pasta foi selecionada
        messagebox.showwarning("Erro", "Selecione uma pasta com arquivos.")  # Alerta o usuário
        return

    if not impressora_selecionada:  # Verifica se uma impressora foi escolhida
        messagebox.showwarning("Erro", "Selecione uma impressora antes de imprimir.")  # Alerta o usuário
        return

    try:
        arquivos = listar_arquivos(caminho)  # Lista os arquivos da pasta
        if not arquivos:  # Verifica se existem arquivos na pasta
            messagebox.showinfo("Aviso", "Nenhum arquivo PDF encontrado na pasta.")  # Alerta o usuário
            return

        progress_bar["maximum"] = len(arquivos)  # Configura o número máximo da barra de progresso
        for idx, arquivo in enumerate(arquivos):
            if os_name == "Windows":  # Se for Windows, usa a função correspondente
                imprimir_windows(arquivo, impressora_selecionada)
            elif os_name == "Linux":  # Se for Linux, usa a função correspondente
                imprimir_linux(arquivo, impressora_selecionada)
            time.sleep(20)  # Pausa de 20 segundos entre as impressões
            progress_bar["value"] = idx + 1  # Atualiza a barra de progresso
            root.update_idletasks()  # Atualiza a interface gráfica
        
        messagebox.showinfo("Sucesso", "Todos os arquivos foram enviados para a impressão!")  # Mensagem de sucesso
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")  # Mensagem de erro

# Função para selecionar a pasta de arquivos
def selecionar_pasta():
    global caminho
    caminho = filedialog.askdirectory()  # Abre uma janela para selecionar pastas
    if caminho:
        pasta_var.set(caminho)  # Atualiza a variável que exibe o caminho na interface

# Função para listar arquivos em uma pasta
def listar_arquivos(pasta):
    """Retorna a lista de arquivos PDF em uma pasta."""
    return [os.path.join(pasta, f) for f in os.listdir(pasta) if f.endswith(".pdf")]

# Identificar o sistema operacional
os_name = platform.system()

# Configurações de layout da janela
janela_largura = 700
janela_altura = 300

# Interface gráfica
root = Tk()
root.title("Impressão de PDFs em Lote")
root.geometry(f"{janela_largura}x{janela_altura}")  # Define o tamanho da janela
root.resizable(False, False)  # Define a janela como não redimensionável

# Variáveis globais para exibir informações na interface
pasta_var = StringVar(value="Nenhuma pasta selecionada")
impressora_var = StringVar(value="Nenhuma impressora selecionada")
caminho = ""
impressora_selecionada = None

# Widgets da interface
Label(root, text="Selecione a pasta com os arquivos PDF:").pack(pady=5)  # Texto para o usuário
Button(root, text="Selecionar Pasta", command=selecionar_pasta).pack(pady=5)  # Botão para selecionar pasta
Label(root, textvariable=pasta_var, wraplength=380).pack(pady=5)  # Exibe o caminho da pasta escolhida

Label(root, text="Selecione a impressora:").pack(pady=5)  # Texto para o usuário
impressoras = listar_impressoras()  # Obtém a lista de impressoras
if impressoras:
    impressora_var.set(impressoras[0])  # Define a primeira impressora como padrão
OptionMenu(root, impressora_var, *impressoras).pack(pady=5)  # Menu para selecionar impressoras

# Botão para confirmar a impressora escolhida
def confirmar_impressora(impressora):
    global impressora_selecionada
    impressora_selecionada = impressora
    messagebox.showinfo("Impressora Selecionada", f"Impressora escolhida: {impressora}")  # Mostra mensagem ao usuário
    impressora_var.set(impressora)

Button(root, text="Aplicar", command=lambda: confirmar_impressora(impressora_var.get())).pack(pady=5)

# Botão para iniciar a impressão
Button(root, text="Imprimir Arquivos", command=imprimir_arquivos, bg="green", fg="white").pack(pady=10)

# Barra de progresso para acompanhar a impressão
progress_bar = ttk.Progressbar(root, orient="horizontal", length=350, mode="determinate")
progress_bar.pack(pady=5)

# Exibição de uma imagem no canto inferior esquerdo
canvas = Canvas(root, width=150, height=150)
canvas.place(x=10, y=150)
try:
    img = PhotoImage(file="rubens.png")  # Substitua pelo caminho da sua imagem
    canvas.create_image(0, 0, anchor="nw", image=img)
except Exception:
    Label(root, text="Imagem não encontrada", fg="red").place(x=10, y=275)

# Loop principal da interface
root.mainloop()