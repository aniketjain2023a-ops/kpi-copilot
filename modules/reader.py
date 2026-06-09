import pandas as pd

def load_report(file_path):
    df = pd.read_excel(file_path)
    return df