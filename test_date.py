import pandas as pd
from datetime import datetime

# Teste para verificar o formato da data
test_dates = [
    "2024-01-15 10:30:00",
    "2024-12-25 14:45:00",
    "2024-03-08 09:15:00"
]

print("Testando formato de data:")
for date_str in test_dates:
    dt = pd.to_datetime(date_str)
    formatted = dt.strftime("%d/%m")
    formatted_full = dt.strftime("%d/%m/%Y")
    print(f"Original: {date_str}")
    print(f"Formato %d/%m: {formatted}")
    print(f"Formato %d/%m/%Y: {formatted_full}")
    print("---")

# Teste com intervalo
data_min = pd.to_datetime("2024-01-15")
data_max = pd.to_datetime("2024-01-20")

if data_min.date() == data_max.date():
    intervalo_data = data_min.strftime("%d/%m/%Y")
else:
    intervalo_data = f"{data_min.strftime('%d/%m')} a {data_max.strftime('%d/%m/%Y')}"

print(f"Intervalo teste: {intervalo_data}")
