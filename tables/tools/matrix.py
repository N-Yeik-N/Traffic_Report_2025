import pandas as pd
import numpy as np

def read_matrix(excel_path):
    df = pd.read_excel(excel_path, header=0, index_col=0)

    destinys = df.columns.tolist()
    origins = df.index.tolist()

    df.replace('-', np.nan, inplace=True)
    df.fillna(0, inplace=True)

    df = df.loc[~(df == 0).all(axis=1)]
    df = df.loc[:, ~(df == 0).all(axis=0)]

    return origins, destinys, df