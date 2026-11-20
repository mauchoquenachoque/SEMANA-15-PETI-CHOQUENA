import pandas as pd
t=pd.read_csv('RG04_tratamiento.csv');t['reduccion']=t.Valor-t['Valor residual'];print(t[['id','Valor','Valor residual','reduccion']].to_string(index=False))
