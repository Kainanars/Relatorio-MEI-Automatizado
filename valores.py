from tkinter import *

# Defina semnotaValue e notaValue como globais para que possam ser acessadas em save
semnotaValue = None
notaValue = None
valores = None

semNotaFinal = 0
comNotaFinal = 0


def set_valores(atividades):
    atividade = 0
    def comercio():
        input_values("Comércio")
        return semNotaFinal, comNotaFinal

    def industria():
        input_values("Indústria")
        return semNotaFinal, comNotaFinal

    def servico():
        input_values("Serviço")
        return semNotaFinal, comNotaFinal

    atividades = set(atividades)
    if {'Comércio', 'Indústria', 'Serviço'} <= atividades:
        atividade = 7
    elif {'Comércio', 'Indústria'} <= atividades:
        atividade = 4
    elif {'Comércio', 'Serviço'} <= atividades:
        atividade = 5
    elif {'Indústria', 'Serviço'} <= atividades:
        atividade = 6
    elif 'Comércio' in atividades:
        atividade = 1
    elif 'Indústria' in atividades:
        atividade = 2
    elif 'Serviço' in atividades:
        atividade = 3

        
    switch = {
        1: comercio,
        2: industria,
        3: servico,
        4: lambda: (comercio(), industria()),
        5: lambda: (comercio(), servico()),
        6: lambda: (industria(), servico()),
        7: lambda: (comercio(), industria(), servico())
    }

    return atividade, switch.get(atividade, lambda: (0, 0))()


def save():
    global semnotaValue, notaValue, valores, semNotaFinal, comNotaFinal
    semnota = float(semnotaValue.get()) if semnotaValue.get() else 0.0
    nota = float(notaValue.get()) if notaValue.get() else 0.0
    semNotaFinal = semnota
    comNotaFinal = nota
    print(semNotaFinal)
    print(comNotaFinal)
    valores.destroy()  # Fechar a janela
    valores.quit()



def input_values(atividade):
    global semnotaValue, notaValue, valores
    valores = Tk()
    valores.title("Relatório MEI Automatizado")

    title = Label(valores, text=f"Atividade de {atividade}")
    title.grid(column=0, row=0)

    semnota_label = Label(valores, text="Receita sem nota")
    semnota_label.grid(column=0, row=1)
    semnotaValue = Entry(valores, width=100)
    semnotaValue.grid(column=0, row=2)

    nota_label = Label(valores, text="Receita com nota")
    nota_label.grid(column=0, row=3)
    notaValue = Entry(valores, width=100)
    notaValue.grid(column=0, row=4)

    botao = Button(valores, text="Salvar", command=save)
    botao.grid(column=0, row=5)

    valores.mainloop()

# testing

# atividade = 1
# atividade, valores = set_valores(atividade)
# print(atividade)
# print(valores[0])
# print(valores[1])
