string = "APSum 8001774-08.2024.8.05.0216"
numeros = ''.join(filter(str.isdigit, string))
formato= f"{numeros[:7]}-{numeros[7:9]}.{numeros[9:13]}.{numeros[13]}.{numeros[14:16]}.{numeros[16:]}"
print(formato)
