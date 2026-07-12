from flask import Flask, render_template, abort, request
import pyodbc
from pathlib import Path
from markupsafe import Markup
import re

app = Flask(__name__)

DB_PATH = Path("Teplo_v31.mdb").resolve()

conn_str = (
    r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
    rf"DBQ={DB_PATH};"
)


def get_conn():
    return pyodbc.connect(conn_str)

def highlight(text, search):

    text = "" if text is None else str(text)

    if not search:
        return text

    pattern = re.compile(re.escape(search), re.IGNORECASE)

    return Markup(
        pattern.sub(
            lambda m: f'<span class="highlight">{m.group(0)}</span>',
            text
        )
    )


@app.route("/")
def index():

    search = request.args.get("q", "").strip()

    conn = get_conn()
    cur = conn.cursor()

    if search:
        rows = cur.execute("""
            SELECT
                ID,
                Name,
                Description
            FROM Object
            WHERE
                Name LIKE ?
                OR Description LIKE ?
                OR CStr(ID) LIKE ?
            ORDER BY Name
        """,
                           f"%{search}%",
                           f"%{search}%",
                           f"%{search}%"
                           ).fetchall()
    else:
        rows = cur.execute("""
            SELECT
                ID,
                Name,
                Description
            FROM Object
            ORDER BY Name
        """).fetchall()

    groups = {
        "ТИТ": [],
        "ТС": [],
        "Измерения": [],
        "Сигналы": [],
        "Прочее": []
    }

    for r in rows:

        item = {
            "id": r.ID,
            "name": r.Name,
            "desc": r.Description or ""
        }

        if r.Name.startswith("ТИТ"):
            groups["ТИТ"].append(item)

        elif r.Name.startswith("ТС"):
            groups["ТС"].append(item)

        elif r.Name.startswith("Измерение"):
            groups["Измерения"].append(item)

        elif r.Name.startswith("Сигнал"):
            groups["Сигналы"].append(item)

        else:
            groups["Прочее"].append(item)

    # Статистика считается ПОСЛЕ заполнения групп
    stats = {
        "Всего объектов": len(rows),
        "ТИТ": len(groups["ТИТ"]),
        "ТС": len(groups["ТС"]),
        "Измерения": len(groups["Измерения"]),
        "Сигналы": len(groups["Сигналы"]),
        "Прочее": len(groups["Прочее"])
    }

    conn.close()

    return render_template(
        "index.html",
        groups=groups,
        search=search,
        stats=stats,
        highlight=highlight
    )


@app.route("/object/<int:obj_id>")
def object_page(obj_id):

    group = request.args.get("group", "")

    conn = get_conn()
    cur = conn.cursor()

    obj = cur.execute("""
        SELECT
            ID,
            Name,
            ParentID,
            Type,
            Description,
            Number,
            Updated,
            RealID,
            UserData
        FROM Object
        WHERE ID=?
    """, obj_id).fetchone()

    if obj is None:
        conn.close()
        abort(404)

    props = cur.execute("""
        SELECT
            Property.ID,
            DicProperty.Name,
            Property.Value,
            Property.TimeStamp,
            Property.Quality
        FROM Property
        LEFT JOIN DicProperty
            ON DicProperty.ID = Property.PropertyID
        WHERE ParentID=?
        ORDER BY DicProperty.Name
    """, obj_id).fetchall()

    conn.close()

    return render_template(
        "object.html",
        obj=obj,
        props=props,
        group=group
    )


if __name__ == "__main__":
    app.run(debug=True)