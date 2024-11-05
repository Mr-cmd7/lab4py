import pandas as pd
import datetime as dt

file = 'data/firms.xlsx'
df = pd.read_excel(file)

current_date = dt.datetime.now().date()
contacts_data = {"Название": [], "Телефон": [], "Адрес": [], "Директор": [], "Дата регистрации": [], "Срок": []}

df['Дата регистрации договора'] = pd.to_datetime(df['Дата регистрации договора'], format='%Y-%m-%d')

for index in range(len(df)):
    firm_name = df.at[index, "Название фирмы"]
    phone_number = df.at[index, "Номер телефона"]
    address = df.at[index, "Адрес"]
    director = df.at[index, "Директор"]
    registration_date = df.at[index, 'Дата регистрации договора'].date()

    if (current_date.month, current_date.day) < (registration_date.month, registration_date.day):
        srok = current_date.year - registration_date.year - 1
    else:
        srok = current_date.year - registration_date.year

    contacts_data["Название"].append(firm_name)
    contacts_data["Телефон"].append(phone_number)
    contacts_data["Адрес"].append(address)
    contacts_data["Директор"].append(director)
    contacts_data["Дата регистрации"].append(registration_date)
    contacts_data["Срок"].append(srok)

contacts_df = pd.DataFrame(contacts_data)

contacts_df = contacts_df.T

with pd.ExcelWriter(file, engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='Фирмы', index=False)
    contacts_df.to_excel(writer, sheet_name='Контакты', index=True, header=False)

    workbook = writer.book
    contacts_worksheet = writer.sheets['Контакты']

    chart = workbook.add_chart({'type': 'column'})
    chart.add_series({
        'name': 'Срок сотрудничества',
        'categories': ['Контакты', 0, 0, 0, len(df) - 1],
        'values': ['Контакты', 5, 1, 5, len(df) - 1],
    })

    chart.set_title({'name': 'Срок сотрудничества с фирмами'})
    chart.set_x_axis({'name': 'Название фирмы'})
    chart.set_y_axis({'name': 'Срок (лет)'})

    contacts_worksheet.insert_chart('L4', chart)
