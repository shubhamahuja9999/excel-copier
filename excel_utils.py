import pandas as pd
import os

def read_excel(path):
    return pd.read_excel(path)

def write_excel(df, path):
    df.to_excel(path, index=False)

def append_to_excel(df, path):
    try:
        if os.path.exists(path):
            try:
                old_df = pd.read_excel(path)
                new_df = pd.concat([old_df, df], ignore_index=True)
            except PermissionError:
                print(f"Error: Cannot access {path}. Please make sure the file is not open in another program.")
                return False
        else:
            new_df = df
            
        try:
            new_df.to_excel(path, index=False)
            return True
        except PermissionError:
            print(f"Error: Cannot write to {path}. Please make sure the file is not open in another program.")
            return False
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return False
