import pandas as pd
pd.options.display.max_columns = None

car_data = pd.read_csv('prueba_xouba.csv')

car_data.columns = car_data.columns.str.rstrip()

print(car_data.head())