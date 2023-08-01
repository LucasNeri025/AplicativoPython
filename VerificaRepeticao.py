import BD
import json

with open('base.json') as file:
   client = json.load(file)

cliente = client["base"]

def verificacao(item):
   for index, i in enumerate(cliente):
      var = f"{index}"
      if item == cliente[index][var]:
         for os in BD.bancoDeDados: 
            if os == item:
               BD.bancoDeDados.pop(BD.bancoDeDados.index(os))
                       
def buscarNan():
  print(BD.bancoDeDados)
   
