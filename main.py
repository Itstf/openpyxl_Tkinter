from tkinter import *
from tkinter import filedialog 
from openpyxl import Workbook, load_workbook

    # lista com todas as abas da planilha --
    # pessoa escolher as abas e colocar em uma lista
    
    # pra cada item dessa lista, um de cada vez ??
    # lista com todas as coluna
   
def browseFiles():
    # montando explorador de arquivos que vai aparecer para o usuário
    filename = filedialog.askopenfilename(initialdir = "/", #caminho pasta que deve ser aberta
                                          title = "Explorer", #titulo do explorador
                                          filetypes = (("Text files", #diferentes arquivos que pode ser filtrado
                                                        "*.txt*"), 
                                                        ("excel",
                                                        "*.xlsx*"),
                                                       ("all files", 
                                                        "*.*"))) 
       
    label_file_explorer.configure(text=filename) #filename = caminho do arquivo
    
    # mostrar as abas que tem no excel
    wb = load_workbook(f'{filename}') #abrir o arquivo
    for sheet in wb: #rodar sobre as abas
        print(sheet)
                                                                                            
window = Tk() #puxar janela
   
# config janela
window.title('File Explorer') #titulo da janela
window.geometry("700x500") #tamanho janela
window.config(background = "green") #cor bg
   
# onde vai aparecer o local do arquivo
label_file_explorer = Label(window,  
                            text = "File path", 
                            width = 100, height = 4,  
                            fg = "blue") 

# botões
button_explore = Button(window,  
                        text = "Browse Files", 
                        command = browseFiles)  
   
button_exit = Button(window,  
                     text = "Exit", 
                     command = exit)  
   
# posições
label_file_explorer.grid(column = 1, row = 1) 
button_explore.grid(column = 1, row = 2) 
button_exit.grid(column = 1,row = 3) 
   
window.mainloop() 