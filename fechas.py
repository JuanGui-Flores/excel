import pandas as pd

def change_date_format(df, column_name):
    # Cambia el formato de fecha en un dataframe Pandas a dd/mm/yyyy.
    df[column_name] = pd.to_datetime(df[column_name]).dt.strftime('%d/%m/%Y')
    return df

def change_estado(df, column_name):
    # Cambia los valores de Estado en un dataframe Pandas.
    df[column_name] = df[column_name].apply(lambda x: 'Funcionalidad Planificada' if x == 'Planificada' else 'Funcionalidad No Planificada')
    return df

def change_tipo_incidencia(df, column_name):
    # Cambia los valores de Tipo de Incidencia en un dataframe Pandas.
    df[column_name] = df[column_name].apply(lambda x: 'Tarea Planificada' if x == 'Tarea Planificada' else 'Tarea No Planificada')
    return df

def save_dataframe_to_excel(df, excel_filename):
    # Guarda un dataframe Pandas en un archivo Excel.
    df.to_excel(excel_filename, index=False)

def main():
    # Leer el archivo Excel
    df = pd.read_excel("C:/Work/jira-search-0b651e7f-5ce9-4912-9e53-c9a4cbd42e77.xlsx")

    # Cambiar el formato de la fecha
    df = change_date_format(df, 'Fecha')

    # Cambiar los valores de Estado
    df = change_estado(df, 'Estado')

    # Cambiar los valores de Tipo de Incidencia
    df = change_tipo_incidencia(df, 'Tipo de Incidencia')

    # Guardar el archivo Excel modificado
    save_dataframe_to_excel(df, "C:/Work/jira-search-0b651e7f-5ce9-4912-9e53-c9a4cbd42e77.xlsx")

if __name__ == '__main__':
    main()
