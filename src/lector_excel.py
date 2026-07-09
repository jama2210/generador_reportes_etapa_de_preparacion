import pandas as pd

def leer_excel(archivo):

    if archivo.name.endswith(".xlsx"):
        df = pd.read_excel(archivo)

    elif archivo.name.endswith(".xls"):
        df = pd.read_excel(archivo)

    else:
        raise Exception("Formato no soportado")

    return df