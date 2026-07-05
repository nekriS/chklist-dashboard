import pandas as pd
import datetime

SUB_HEAD = [
    ["Компонент",               "Схемсимвол",   "БД+Символ",    "Посадочное",   "Перевод"   ]
]

REFRESH = '<meta http-equiv="refresh" content="60">'

def highlight_by_value(row, config, type="main"):

    yellow = f"background-color: #{config['COLORS']['YELLOW']}"
    red = f"background-color: #{config['COLORS']['RED']}"
    blue = f"background-color: #{config['COLORS']['BLUE']}"
    green = f"background-color: #{config['COLORS']['GREEN']}"

    if type == "main":
        color = ''
        val = row.iloc[3]
        match val:
            case "Новый":
                color = yellow
            case "Идет проверка":
                color = red
            case "К переводу":
                color = blue
            case "Готов номинально":
                color = green
            case "Готов":
                color = green
            case "Готов вручную":
                color = green
            case _:
                color = yellow

        return [color] * len(row)
    else:
        color = ['', '']
        for i in range(2, 5):
            val = row.iloc[i]
            match val:
                case "Не проверено":
                    color.append(yellow)
                case "Нет":
                    color.append(red)
                case "Да":
                    color.append(green)

        return color



def drawHDashboard(config, dataframe, subtables: dict):

    pathDashboardFile = f"{config["GENERAL"]['DEFAULT_PATH']}dashboard.html"
    pathSubFile = f"{config["GENERAL"]['DEFAULT_PATH']}{config["GENERAL"]['NAME_FOLDER_WEB']}/"

    for dataRow in dataframe.values:
        project = dataRow[1]

        subdf = pd.DataFrame(subtables[project])

        subdf.columns = SUB_HEAD[0]

        styled_subdf = subdf.style.apply(highlight_by_value, axis=1, config=config, type="sub")\
        .set_table_styles([
            {'selector': '', 'props': [('border', '2px solid black'), 
                                   ('border-collapse', 'collapse')]},

            {'selector': 'th', 'props': [('background-color', 'white'), 
                                        ('color', 'black'), 
                                        ('border', '1px solid black'),
                                        ('font-weight', 'bold'),
                                        ('font-family', 'Arial'),
                                        ('padding', '10px')]},

            {'selector': 'td', 'props': [('border', '1px solid black'), 
                                        ('padding', '8px'),
                                        ('font-family', 'Arial'),
                                        ('text-align', 'center')]}
        ])\
        .set_properties(subset=['Компонент'], **{'width': '140px'})\
        .set_properties(subset=['БД+Символ'], **{'width': '140px'})\
        .set_properties(subset=['Посадочное'], **{'width': '140px'})\
        .set_properties(subset=['Перевод'], **{'width': '140px'})\
        .hide(axis='index')

        link_html = f'<a href="../" style="display: block; margin: 15px; font-family: Arial;">← Вернуться назад</a>\
            <a style="display: block; font-family: Arial;">Проект: {project}</a>'
        final_html = REFRESH + link_html + styled_subdf.to_html(escape=False, index=False)

        with open(pathSubFile + project + ".html", "w", encoding="utf-8") as f:
            f.write(final_html)

    dataframe['Имя проекта'] = dataframe['Имя проекта'].apply(lambda x: f'<a href="web/{x}.html">{x}</a>')

    styled_df = dataframe.style.apply(highlight_by_value, axis=1, config=config)\
        .set_table_styles([
            {'selector': '', 'props': [('border', '2px solid black'), 
                                   ('border-collapse', 'collapse')]},
   
            {'selector': 'th', 'props': [('background-color', 'white'), 
                                        ('color', 'black'), 
                                        ('border', '1px solid black'),
                                        ('font-weight', 'bold'),
                                        ('font-family', 'Arial'),
                                        ('padding', '10px')]},

            {'selector': 'td', 'props': [('border', '1px solid black'), 
                                        ('padding', '8px'),
                                        ('font-family', 'Arial'),
                                        ('text-align', 'center')]}
        ])\
        .hide(axis='index')
    
    date_html = f'<a style="display: block; font-family: Arial;">Последнее обновление: {datetime.datetime.now()}</a>'

    console_html = """"""
    
    
    final_html = REFRESH + date_html + styled_df.to_html(escape=False, index=False) + console_html

    with open(pathDashboardFile, "w", encoding="utf-8") as f:
        f.write(final_html)


    


    