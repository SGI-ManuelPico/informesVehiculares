import pandas as pd


class Indicadores():

    def __init__(self):
        pass

    def calcular_EJL(df_diario):
        # Asegurarse de que la columna 'FECHA' es de tipo datetime en el DataFrame original
        df_diario['FECHA'] = pd.to_datetime(df_diario['FECHA'], format="%d/%m/%Y")

        # Crear una columna para el mes y año
        df_diario['MES'] = df_diario['FECHA'].dt.to_period('M')

        # Calcular los acumulados mensuales
        df_acumulados = df_diario.groupby('MES').sum(numeric_only=True).reset_index()

        # Crear un DataFrame con todos los meses del año
        all_months = pd.date_range(start='2024-01-01', end='2024-12-31', freq='M').to_period('M')
        df_all_months = pd.DataFrame(all_months, columns=['MES'])

        # Unir con el DataFrame de acumulados mensuales para asegurarse de que todos los meses estén presentes
        df_acumulados = pd.merge(df_all_months, df_acumulados, on='MES', how='left')

        # Rellenar los NaN con ceros y asegurarse de que el tipo de datos sea correcto
        df_acumulados.fillna(0, inplace=True)

        # Cambiar el formato del periodo a nombre de mes en español
        df_acumulados['MES'] = df_acumulados['MES'].dt.strftime('%B').str.capitalize()

        # Calcular EJL
        df_EJL = pd.DataFrame({
            'MES': df_acumulados['MES'],
            'EJD': df_acumulados['EXCESOS VELOCIDAD'],
            'SDT': df_acumulados['DÍA TRABAJADO'],
            'EJL': (df_acumulados['EXCESOS VELOCIDAD'] / df_acumulados['DÍA TRABAJADO']) * 100
        })

        # Redondear a 2 decimales
        df_EJL = df_EJL.round(2)

        # Transponer el DataFrame
        df_EJL = df_EJL.set_index('MES').transpose()

        return df_EJL

    def calcular_GVE(df_diario, df_hist):
        # Asegurarse de que la columna 'FECHA' es de tipo datetime en el DataFrame original
        df_diario['FECHA'] = pd.to_datetime(df_diario['FECHA'], format="%d/%m/%Y")

        # Crear una columna para el mes y año
        df_diario['MES'] = df_diario['FECHA'].dt.to_period('M')

        # Calcular los acumulados mensuales
        df_acumulados = df_diario.groupby('MES').sum(numeric_only=True).reset_index()

        # Calcular el valor máximo de "DÍA TRABAJADO" por mes
        df_max_dia_trabajado = df_diario.groupby('MES')['DÍA TRABAJADO'].max().reset_index()

        # Crear un DataFrame con todos los meses del año
        all_months = pd.date_range(start='2024-01-01', end='2024-12-31', freq='M').to_period('M')
        df_all_months = pd.DataFrame(all_months, columns=['MES'])

        # Unir los DataFrames de acumulados mensuales y máximos de día trabajado para asegurarse de que todos los meses estén presentes
        df_acumulados = pd.merge(df_all_months, df_acumulados, on='MES', how='left')
        df_max_dia_trabajado = pd.merge(df_all_months, df_max_dia_trabajado, on='MES', how='left')

        # Rellenar los NaN con ceros y asegurarse de que el tipo de datos sea correcto
        df_acumulados.fillna(0, inplace=True)
        df_max_dia_trabajado.fillna(0, inplace=True)

        # Cambiar el formato del periodo a nombre de mes en español
        df_acumulados['MES'] = df_acumulados['MES'].dt.strftime('%B').str.capitalize()
        df_max_dia_trabajado['MES'] = df_max_dia_trabajado['MES'].dt.strftime('%B').str.capitalize()

        # Calcular VIP y VLD
        vip_value = len(df_hist['placa'].unique())  # VIP es constante para todos los meses
        vld_values = df_max_dia_trabajado['DÍA TRABAJADO'].tolist()  # VLD es el valor máximo de DÍA TRABAJADO para cada mes

        # Calcular GVE
        gve_values = [(vip_value / vld) * 100 if vld != 0 else 0 for vld in vld_values]

        # Crear el DataFrame de GVE
        df_GVE = pd.DataFrame({
            'MES': df_acumulados['MES'],
            'VIP': [vip_value] * len(df_acumulados),
            'VLD': vld_values,
            'GVE': gve_values
        })

        # Redondear a 2 decimales
        df_GVE = df_GVE.round(2)

        # Transponer el DataFrame
        df_GVE = df_GVE.set_index('MES').transpose()

        return df_GVE

    def calcular_ELVL(df_diario):
        # Asegurarse de que la columna 'FECHA' es de tipo datetime en el DataFrame original
        df_diario['FECHA'] = pd.to_datetime(df_diario['FECHA'], format="%d/%m/%Y")

        # Crear una columna para el mes y año
        df_diario['MES'] = df_diario['FECHA'].dt.to_period('M')

        # Calcular los acumulados mensuales
        df_acumulados = df_diario.groupby('MES').sum(numeric_only=True).reset_index()

        # Crear un DataFrame con todos los meses del año
        all_months = pd.date_range(start='2024-01-01', end='2024-12-31', freq='M').to_period('M')
        df_all_months = pd.DataFrame(all_months, columns=['MES'])

        # Unir con el DataFrame de acumulados mensuales para asegurarse de que todos los meses estén presentes
        df_acumulados = pd.merge(df_all_months, df_acumulados, on='MES', how='left')

        # Rellenar los NaN con ceros y asegurarse de que el tipo de datos sea correcto
        df_acumulados.fillna(0, inplace=True)

        # Cambiar el formato del periodo a nombre de mes en español
        df_acumulados['MES'] = df_acumulados['MES'].dt.strftime('%B').str.capitalize()

        # Calcular ELVL
        df_ELVL = pd.DataFrame({
            'MES': df_acumulados['MES'],
            'DLEV': df_acumulados['EXCESOS VELOCIDAD'],
            'TDL': df_acumulados['DESPLAZAMIENTOS'],
            'ELVL': (df_acumulados['EXCESOS VELOCIDAD'] / df_acumulados['DESPLAZAMIENTOS']) * 100
        })

        # Redondear a 2 decimales
        df_ELVL = df_ELVL.round(2)

        # Transponer el DataFrame
        df_ELVL = df_ELVL.set_index('MES').transpose()

        return df_ELVL

    def calcular_IDP(df_diario):
        # Asegurarse de que la columna 'FECHA' es de tipo datetime en el DataFrame original
        df_diario['FECHA'] = pd.to_datetime(df_diario['FECHA'], format="%d/%m/%Y")

        # Crear una columna para el mes y año
        df_diario['MES'] = df_diario['FECHA'].dt.to_period('M')

        # Calcular los acumulados mensuales
        df_acumulados = df_diario.groupby('MES').sum(numeric_only=True).reset_index()

        # Crear un DataFrame con todos los meses del año
        all_months = pd.date_range(start='2024-01-01', end='2024-12-31', freq='M').to_period('M')
        df_all_months = pd.DataFrame(all_months, columns=['MES'])

        # Unir con el DataFrame de acumulados mensuales para asegurarse de que todos los meses estén presentes
        df_acumulados = pd.merge(df_all_months, df_acumulados, on='MES', how='left')

        # Rellenar los NaN con ceros y asegurarse de que el tipo de datos sea correcto
        df_acumulados.fillna(0, inplace=True)

        # Cambiar el formato del periodo a nombre de mes en español
        df_acumulados['MES'] = df_acumulados['MES'].dt.strftime('%B').str.capitalize()

        # Calcular IDP (Placeholder, ajusta según la fórmula correcta)
        df_IDP = pd.DataFrame({
            'MES': df_acumulados['MES'],
            'DPV': df_acumulados['PREOPERACIONAL'],
            'DTR': df_acumulados['DÍA TRABAJADO'],
            'IDP': (df_acumulados['PREOPERACIONAL'] / df_acumulados['DÍA TRABAJADO']) * 100
        })

        # Redondear a 2 decimales
        df_IDP = df_IDP.round(2)

        # Transponer el DataFrame
        df_IDP = df_IDP.set_index('MES').transpose()

        return df_IDP