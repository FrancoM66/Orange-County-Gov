import pandas as pd


df = pd.read_csv('test.csv',skiprows=9)
df2 = pd.read_csv("O:\Field Services Division\Field Support Center\Project Acceptance\99999 - project\Excel\Development Pressure.csv")


df3 = pd.concat([df2["Location"],df["ID Number"].dropna()])

df3.to_csv('file_name.csv', index=False)