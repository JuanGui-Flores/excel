import pandas as pd

# lista de fechas
fechas = ['2022-01-01', '2022-02-01', '2022-03-01', '2022-04-01', '2022-05-01']

# crear un DataFrame de pandas con las fechas
df = pd.DataFrame({'Fechas': fechas})

# guardar el DataFrame en un archivo de Excel
df.to_excel('fechas.xlsx', index=False)
