import pandas as pd

NO_RESPONDIDO = "(No respondido en el formulario)"

def limpiar_valor(valor):
    if pd.isna(valor):
        return NO_RESPONDIDO

    valor = str(valor).strip()

    if valor == "":
        return NO_RESPONDIDO

    return valor