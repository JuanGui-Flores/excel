import pandas as pd

def cambiar_formato_fecha(fecha):
    return fecha.strftime('%d-%m-%Y')

def cambiar_valores(valor):
    if valor == 'A':
        return 'Funcionalidad Planificada'
    elif valor == 'B':
        return 'Tarea Planificada'
    elif valor == 'C':
        return 'Tarea No Planificada'
    else:
        return valor

df = pd.read_excel("C:\Work\jira-search-0b651e7f-5ce9-4912-9e53-c9a4cbd42e77.xlsx")

df['fecha'] = df['fecha'].apply(cambiar_formato_fecha)
df['estado'] = df['estado'].apply(cambiar_valores)
df['tipo_incidencia'] = df['tipo_incidencia'].apply(cambiar_valores)

df.to_excel("C:\Work\jira-search-0b651e7f-5ce9-4912-9e53-c9a4cbd42e77.xlsx", index=False)
