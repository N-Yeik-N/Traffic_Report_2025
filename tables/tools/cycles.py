import pandas as pd

def get_dates_cycles(codes: str) -> list:
    DATA_PATH = r".\data\Datos de Ciclos.xlsx"
    df = pd.read_excel(DATA_PATH, sheet_name="Hoja1", header=0, usecols="A:E")
    df['Fecha 1'] = df['DIA 1'].apply(lambda x: ' '.join(x.split()[1:]))
    df['Fecha 2'] = df['DIA 2'].apply(lambda x: ' '.join(x.split()[1:]))
    
    df['Fecha 1'] = pd.to_datetime(df['Fecha 1'], format="%d/%m/%y")
    df['Fecha 2'] = pd.to_datetime(df['Fecha 2'], format="%d/%m/%y")

    df['Dia Atipico'] = ""
    df['Dia Tipico'] = ""

    for index, row in df.iterrows():
        if not pd.isnull(row['Fecha 1']) and row['Fecha 1'].day_name() in ['Saturday', 'Sunday']:
            df.at[index, 'Dia Atipico'] = row['Fecha 1'].strftime('%d/%m/%Y')
        else:
            df.at[index, 'Dia Tipico'] = row['Fecha 1'].strftime('%d/%m/%Y') if not pd.isnull(row['Fecha 1']) else ""
            
        if not pd.isnull(row['Fecha 2']) and row['Fecha 2'].day_name() in ['Saturday', 'Sunday']:
            df.at[index, 'Dia Atipico'] = row['Fecha 2'].strftime('%d/%m/%Y')
        else:
            df.at[index, 'Dia Tipico'] = row['Fecha 2'].strftime('%d/%m/%Y') if not pd.isnull(row['Fecha 2']) else ""

    df = df.drop(['Item','Intersección','DIA 1','DIA 2','Fecha 1','Fecha 2'], axis=1)
    
    filtered_df = df[df['Código'].isin(codes)].reset_index(drop=True)

    return filtered_df