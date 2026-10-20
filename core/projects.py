
import sqlite3
from datetime import datetime as time
from datetime import timedelta as delta

class projectsObject:

    def __init__(self, db_path="projects.db", log=None):
        self.conn = sqlite3.connect(db_path, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self.cursor = self.conn.cursor()
        self.logger = log
        self.prefix = "[P]"
        self.table_name = "projects"
        self._initialization()

    def _log(self, message):
        if self.logger == None:
            print(f"{self.prefix} [{time.now()}] > {message}")
        else:
            self.logger(f"{self.prefix} {message}")

    def _initialization(self):
        self.cursor.execute(f"""
            CREATE TABLE IF NOT EXISTS {self.table_name} (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_name TEXT,
                path TEXT,
                developer TEXT,
                state INTEGER,
                quantity INTEGER,
                bd INTEGER,
                fp INTEGER,
                noTMP INTEGER,
                visible INTEGER,
                is_read INTEGER DEFAULT 0,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(project_name)
            )
        """)
        self.conn.commit()

    def process_batch(self, batch: list[dict]):
        try:
            actual_changes_count = 0
            for item in batch:
                action = item.get("action", "save")

                if action == "save":
                    # Безопасный SQL с параметрами ? и конструкцией Upsert
                    self.cursor.execute(f"""
                        INSERT INTO {self.table_name} (
                            project_name, path, developer, state, quantity, bd, fp, noTMP, visible, is_read, updated_at
                        ) VALUES (
                            ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, STRFTIME('%Y-%m-%d %H:%M:%S', 'now', 'localtime')
                        )
                        ON CONFLICT(project_name) DO UPDATE SET
                            path = excluded.path,
                            developer = excluded.developer,
                            state = excluded.state,
                            quantity = excluded.quantity,
                            bd = excluded.bd,
                            fp = excluded.fp,
                            noTMP = excluded.noTMP,
                            visible = excluded.visible,
                            is_read = {self.table_name}.is_read,
                            updated_at = excluded.updated_at
                        WHERE 
                            {self.table_name}.path != excluded.path OR
                            {self.table_name}.developer != excluded.developer OR
                            {self.table_name}.state != excluded.state OR
                            {self.table_name}.quantity != excluded.quantity OR
                            {self.table_name}.bd != excluded.bd OR
                            {self.table_name}.fp != excluded.fp OR
                            {self.table_name}.noTMP != excluded.noTMP OR
                            {self.table_name}.visible != excluded.visible
                    """, (
                        item["project_name"],
                        item["path"],
                        item["developer"],
                        item["state"],
                        item["quantity"],
                        item["bd"],
                        item["fp"],
                        item["noTMP"],
                        item["visible"]
                    ))

                elif action == "delete":
                    self.cursor.execute(f"""
                        DELETE FROM {self.table_name} 
                        WHERE id = ?
                    """, (item["id"],))

                if self.cursor.rowcount > 0:
                    actual_changes_count += 1

            self.conn.commit()
            if actual_changes_count > 0:
                self._log(f"A batch of {actual_changes_count} items has been successfully processed and updated.")

        except sqlite3.Error as e:
            self.conn.rollback()
            self._log(f"Error during transaction execution: {e}")
            raise e

    def get_all_active(self) -> list[dict]:
        try:
            self.cursor.execute(f"SELECT * FROM {self.table_name}")
            return [dict(row) for row in self.cursor.fetchall()]
        except sqlite3.Error as e:
            self._log(f"Error fetching all active projects: {e}")
            return []
        
    def get_project(self, project_name) -> dict:
        try:
            self.cursor.execute(f"SELECT * FROM {self.table_name} WHERE project_name = ?", (project_name,))
            return [dict(row) for row in self.cursor.fetchall()][0]
        except sqlite3.Error as e:
            self._log(f"Error fetching project: {e}")
            return []

    def get_project_by_id(self, id: int) -> dict:
        try:
            self.cursor.execute(f"SELECT * FROM {self.table_name} WHERE id = ?", (id,))
            return [dict(row) for row in self.cursor.fetchall()][0]
        except sqlite3.Error as e:
            self._log(f"Error fetching project: {e}")
            return []

    def names(self):
        names = []
        for project in self.get_all_active():
            names.append(project["project_name"])
        return names

    def add_new_path(self, project_name, path):
        project = self.get_project(project_name)
        paths = project["path"].split("|")
        paths.append(path)
        project["path"] = "|".join(paths)
        self.process_batch([project])

    def get_last_path(self, project_name):
        project = self.get_project(project_name)
        paths = project["path"].split("|")
        return paths[-1]
    
    def close(self):
        self.conn.close()