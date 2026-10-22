from datetime import datetime as time
import os
import shutil
import pandas as pd
from pathlib import Path
from system import log
from .combine import combine_checklists
from addons.avatar import create_number_image
from .states import check_state
from worksheets.services import drawXDashboard
from web.services import drawHDashboard
from pyodbc import connect
import formulas

def get_dashboard_html(dashboard):

    projects = dashboard.projects.get_all_active()

    if len(projects) > 1:

        data_dataframe = getDataframe(projects)

        df2 = data_dataframe.iloc[:, [0, 1, 5, 6, 7]]

        table_html = "Таблица состояний<br><br>" + df2.to_html(index=False) + f"<br>Актуально на: {time.now().strftime('%Y.%m.%d %H:%M')}"

        table_html = table_html.replace('<table border="1" class="dataframe">', '<table class="dataframe" style="border: 1px solid rgb(155, 155, 155); overflow-wrap: break-word; border-collapse: collapse;"> <col span="5"> <col style="width: 100px;"> <col style="width: 0%;"> <col style="width: 100px;"> <col style="width: 0%;"> <col style="width: 100px;">')
        table_html = table_html.replace('<th>', '<th style="border: 1px solid rgb(155, 155, 155); overflow-wrap: break-word; border-collapse: collapse;">')
        table_html = table_html.replace('<tr style="text-align: right;">','<tr style="text-align: center;">')
        table_html = table_html.replace('<tr>','<tr style="text-align: center;">')
        table_html = table_html.replace('<td>','<td style="border: 1px solid rgb(155, 155, 155); overflow-wrap: break-word; border-collapse: collapse;">')
    else:
        table_html = "Нет текущих проектов."

    return table_html

def get_tmp_parts_from_db(config):

    try:

        DRIVER = "{" + config["DB"]["DRIVER"] + "}"
        SQL_CONNECTION_STRING = f'Driver={DRIVER};Server={config["DB"]["SERVER"]};Database={config["DB"]["DATABASE"]};Uid={config["DB"]["USER"]};Pwd={config["DB"]["PASSWORD"]};'
        conn = connect(SQL_CONNECTION_STRING)

        # GETTING TABLES NAMES
        SQL_QUERY = """
        SELECT name
        FROM sys.tables;
        """

        cursor = conn.cursor()
        cursor.execute(SQL_QUERY)
        #print(1)
        records = cursor.fetchall()

        tables = []
        for r in records:
            line = []
            for m in r:
                line.append(m)
            tables.append(line[0])

        components = []

        for table_name in tables:

            if table_name != "TMPPRTS":
                SQL_QUERY = f"""
                SELECT [PART_NUMBER]
                FROM [{table_name}];
                """

                cursor = conn.cursor()
                cursor.execute(SQL_QUERY)
                records = cursor.fetchall()


                for r in records:
                    line = []
                    for m in r:
                        line.append(m)
                    if "TMP" in line[0]:
                        components.append(line[0])

        SQL_QUERY = """
        SELECT *
        FROM TMPPRTS;
        """

        cursor = conn.cursor()
        cursor.execute(SQL_QUERY)
        records = cursor.fetchall()
        table = []
        del_components = []
        for r in records:
            line = []
            for m in r:
                line.append(m)
            table.append(line)
            if line[0] not in components and line[1] == None:
                del_components.append(line[0])

        df = pd.DataFrame(table, index=None, columns=["TMP_NAME", "NAME", "AUTHOR", "DATE"])

        df2 = pd.DataFrame(del_components, index=None, columns=["TMP_NAME"])

        df.to_csv(f"{config["GENERAL"]['DEFAULT_PATH']}{config["GENERAL"]['NAME_FOLDER_DATA']}/{config["GENERAL"]['NAME_FILE_BD']}", index=False)
        df2.to_csv(f"{config["GENERAL"]['DEFAULT_PATH']}{config["GENERAL"]['NAME_FOLDER_DATA']}/{config["GENERAL"]['NAME_FILE_BD_DEL']}", index=False)
        log(f"BD with TMP parts was updated successful.")


    except Exception as e:
        log(f"ERROR: {e}")

def add_new_checklist(projects, path):
    try:
        file_head = pd.read_excel(f"{path}", header=None, nrows=4, usecols="B").fillna("_NONE").values
        table = pd.read_excel(f"{path}", header=None, usecols="A",skiprows=7).fillna("_NONE").values
        project_name = file_head[0][0]
        developer = file_head[1][0]

        projects.process_batch([{
            "action": "save",
            "project_name": project_name,
            "path": path,
            "developer": developer,
            "state": 0,
            "quantity": len(table),
            "bd": 0,
            "fp": 0,
            "noTMP": 0,
            "visible": 1
        }])

        log(f"Project {project_name} was added!")
    except Exception as e:
        log(f"ERROR: {e}")




def get_cell_values_by_name(file_path: str, target_sheet: str, target_cells: list[str]) -> list[str]:

    abs_path = os.path.abspath(file_path)
    xl_model = formulas.ExcelModel().loads(abs_path).finish()

    res_solution = xl_model.calculate()

    target_suffixies = [f"]{target_sheet.upper()}!{target_cell.strip().upper()}" for target_cell in target_cells]

    result = []
    for target_suffix in target_suffixies:
        for key, cell_value in res_solution.items():
            clean_key = key.replace("'", "").replace('"', "").upper()

            if clean_key.endswith(target_suffix):
                actual_value = getattr(cell_value, 'value', cell_value)
                if isinstance(actual_value, (list, bytes, str)) and hasattr(actual_value, 'tolist'):
                    actual_value = actual_value.tolist()
                result.append(actual_value.item())

    return result

def force_formulas(file_name, df, start=8):
    excel_file = pd.ExcelFile(file_name)
    first_sheet_name = excel_file.sheet_names[0]
    columns = ["C", "E", "J"]
    quantity_rows = len(df)
    values = get_cell_values_by_name(file_name, 
                                     first_sheet_name, 
                                     [f"{column}{num}" for num in range(start, start+quantity_rows+1) for column in columns])

    j = 0
    for idx_pos, (index, row) in enumerate(df.iterrows()):
        for column in [1, 3, 5]:
            df.iat[idx_pos, column] = values[j]
            j += 1

    return df

def getComponents(config, project: dict):
    components = {}
    try:
        TMP_PARTS = pd.read_csv(f"{config["GENERAL"]['DEFAULT_PATH']}{config["GENERAL"]['NAME_FOLDER_DATA']}/{config["GENERAL"]['NAME_FILE_BD']}").dropna(subset=["NAME"])['TMP_NAME'].values
    except:
        TMP_PARTS = []
    try:
        DEL_TMP_PARTS = pd.read_csv(f"{config["GENERAL"]['DEFAULT_PATH']}{config["GENERAL"]['NAME_FOLDER_DATA']}/{config["GENERAL"]['NAME_FILE_BD_DEL']}")['TMP_NAME'].values
    except:
        DEL_TMP_PARTS = []

    for path in [project["path"].split("|")[-1]]:

        with pd.option_context('future.no_silent_downcasting', True):
            df = pd.read_excel(f"{path}", header=None, names=["B", "C", "D", "E", "F", "J", "K"], usecols="B, C, D, E, F, J, K",skiprows=7)
            column_order = ["B", "C", "D", "E", "F", "J", "K"]
            df = df[column_order].fillna('_NONE')
            df = df.replace(' ', '_NONE')

            if df.iloc[0, 1] == "_NONE":
                df = force_formulas(path, df)

            df = df.replace(to_replace={'Не проверено': 0, 'Да': 1, 'Нет': -1})
            table = df.values


        for lines in table:

            if lines[1] == -1 or lines[3] == -1:
                lines13 = -1
            elif lines[1] == 1 and lines[3] == 1:
                lines13 = 1
            else:
                lines13 = 0

            if lines[0] not in components:
                components[lines[0]] = [lines13, lines[5], lines[2], lines[4], lines[6]]
            else:
                if lines13 != 0:
                    components[lines[0]][0] = lines13
                if lines[5] != 0 and lines[5] != "_NONE":
                    components[lines[0]][1] = lines[5]
                if lines[2] != "_NONE":
                    components[lines[0]][2] = lines[2]
                if lines[4] != "_NONE":
                    components[lines[0]][3] = lines[4]
                if lines[6] != "_NONE":
                    components[lines[0]][4] = lines[6]

    for component in components.keys():
        if component in DEL_TMP_PARTS: 
            components[component].append(1)
        elif component not in TMP_PARTS:
            components[component].append(2)
        else:
            components[component].append(0)
    return components


def sync_status(dashboard):

    projects = dashboard.projects.get_all_active()

    for project in projects:

        components = getComponents(dashboard.config, project)
        
        yes_bd_schem = 0
        no_bd_schem = 0
        pcb_yes = 0
        pcb_no = 0
        noTMP = 0
        for comp in components:
            if components[comp][0] == 1:
                yes_bd_schem += 1
            elif components[comp][0] == -1:
                no_bd_schem += 1
            if components[comp][1] == 1:
                pcb_yes += 1
            elif components[comp][1] == -1:
                pcb_no += 1
            if components[comp][5] < 2:
                noTMP += 1

        project["noTMP"] = noTMP
        project["fp"] = pcb_yes
        project["bd"] = yes_bd_schem
        project["quantity"] = len(components)

        if noTMP == len(components) and (yes_bd_schem != len(components) or pcb_yes != len(components)):
            project["state"] = 9


        check_state(dashboard, project, components)

    dashboard.projects.process_batch(projects)

def getStateByNumber(stateNumber: int) -> str:
    match stateNumber:
        case 0 | 1 | 2:
            return "Новый"
        case 3 | 4 | 401 | 5:
            return "Идет проверка"
        case 6:
            return "К переводу"
        case 8:
            return "Готов вручную"
        case 9:
            return "Готов номинально"
        case 10:
            return "Готов"
        case _:
            return ""

def getSubTable(config, project: dict) -> list:
    components = getComponents(config, project)
    table = []
    for component in components:
        line = []
        line.append(component)
        line.append(components[component][2])
        if components[component][0] == 1:
            line.append('Да')
        elif components[component][0] == -1:
            line.append('Нет')
        else:
            line.append('Не проверено')
        if components[component][1] == 1:
            line.append('Да')
        elif components[component][1] == -1:
            line.append('Нет')
        else:
            line.append('Не проверено')
        if components[component][5] == 0:
            line.append('Да')
        elif components[component][5] == 1:
            line.append('Удален')
        else:
            line.append('Нет')
        table.append(line)

    return table

def getSubtables(config, projects: list):
    subtables = {}
    for project in projects:
        subtable = getSubTable(config, project)
        subtables[project["project_name"]] = subtable
    return subtables

def getDataframe(projects: list):
    dataList = []
    for project in projects:
        if not project["visible"]:
            continue
        
        row_data = []
        row_data.append(project["id"])
        row_data.append(project["project_name"])
        row_data.append(project["developer"])
        row_data.append(getStateByNumber(int(project["state"])))
        row_data.append(time.strptime(project["updated_at"], "%Y-%m-%d %H:%M:%S"))
        row_data.append(f"{project["bd"]:>3} / {project["quantity"]:<3}")
        row_data.append(f"{project["fp"]:>3} / {project["quantity"]:<3}")
        row_data.append(f"{project["noTMP"]:>3} / {project["quantity"]:<3}")
        dataList.append(row_data)

    data_df = pd.DataFrame(dataList)
    status_order = ["Новый", "Идет проверка", "К переводу", "Готов", "Готов номинально", "Готов вручную"][::-1]
    data_df[3] = pd.Categorical(data_df[3], categories=status_order, ordered=True)
    projectsDataframe = data_df.sort_values(3).reset_index(drop=True)
    projectsDataframe.columns = ['ID', 'Имя проекта', 'Разработчик', 'Статус', 'Последнее обновление', 'БД+Символ', 'Посадочное', 'Переведен в постоянные']

    return projectsDataframe

def updateDashboard(dashboard):

    projects = dashboard.projects.get_all_active()
    dataframe = getDataframe(projects)
    subtables = getSubtables(dashboard.config, projects)

    #drawXDashboard(config, dataframe, subtables)
    drawHDashboard(dashboard.config, dataframe, subtables)

def check_function(dashboard):

    path_upload_folder = f"{dashboard.config["GENERAL"]['DEFAULT_PATH']}{dashboard.config["GENERAL"]['NAME_FOLDER_UPLOADS']}"
    path_work_folder = f"{dashboard.config["GENERAL"]['DEFAULT_PATH']}{dashboard.config["GENERAL"]['NAME_FOLDER_INWORK']}"
    today_date = time.now().strftime('%Y_%m_%d')
    current_time = time.now().strftime('%H_%M_%S')
    upload_files = os.listdir(path_upload_folder)
    projects_count_old = len(dashboard.projects.get_all_active())

    if len(upload_files) > 0:
        for file in upload_files:
            try:     
                file_head = pd.read_excel(f"{path_upload_folder}/{file}", header=None, nrows=4, usecols="B").fillna("_NONE").values
                project_name = file_head[0][0]
                new_file_path = f"{path_work_folder}/{project_name}_{today_date}_{current_time}.xlsx".replace(" ", "_")
                shutil.move(Path(f"{path_upload_folder}/{file}"), Path(new_file_path))

                if not(project_name in dashboard.projects.names()):
                    add_new_checklist(dashboard.projects, new_file_path)
                else:
                    last_file = dashboard.projects.get_last_path(project_name)
                    dashboard.projects.add_new_path(project_name, new_file_path)

                    corr_file_path = f"{path_work_folder}/{project_name}_{today_date}_{current_time}_corr.xlsx".replace(" ", "_")
                    combine_checklists(last_file, new_file_path, corr_file_path)
                    dashboard.projects.add_new_path(project_name, corr_file_path)
                    
                    log(f"New PATH in {project_name} was added!")
            except Exception as e:
                log(f"[WARN] Failed to check upload files: {e}")

    try:
        projects_count = len(dashboard.projects.get_all_active())
        if projects_count != projects_count_old:
            dashboard.messanger.set_avatar(create_number_image(projects_count))
    except Exception as e:
        log(f"ERROR1: {e}")


    try:
        sync_status(dashboard)
    except Exception as e:
        log(f"[WARN] Failed to check status: {e}")
    try:
        updateDashboard(dashboard)
    except Exception as e:
        log(f"[WARN] Failed to update dashboard: {e}")

    