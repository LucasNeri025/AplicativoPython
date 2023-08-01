import BD

def verifica():
    for i in range(len(BD.bancoDeDados)):
        if not BD.bancoDeDados[i]in BD.bancoDeDados[i+1:]:
            BD.bancoDeDados2.append(BD.bancoDeDados[i])

def mostra():
    for i,elemento in enumerate(BD.bancoDeDados2):
        print(i,elemento)
