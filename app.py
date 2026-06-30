from fastapi import FastAPI, Request
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles
from fastapi.responses import JSONResponse
from dotenv import load_dotenv
load_dotenv()
import os
import psycopg2
import psycopg2.extras

app = FastAPI()

app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")


def get_db():
    return psycopg2.connect(os.getenv("DATABASE_STRING"))


@app.get("/")
async def home(request: Request):
    return templates.TemplateResponse("base.html", {"request": request})


@app.get("/rates")
async def rates(request: Request):
    conn = get_db()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    cur.execute("SELECT * FROM materials ORDER BY is_active DESC NULLS LAST, created_at DESC;")
    mat_rows = cur.fetchall()
    cur.execute("SELECT * FROM burden_rates ORDER BY is_active DESC NULLS LAST, created_at DESC;")
    rates_rows = cur.fetchall()
    cur.close()
    conn.close()
    return templates.TemplateResponse("index.html", {"request": request, "materials": mat_rows, "burden_rates": rates_rows})


@app.post("/materials/{material_id}/select")
async def select_material(material_id: int):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("UPDATE materials SET is_active = false")
    cur.execute("UPDATE materials SET is_active = true WHERE id = %s", (material_id,))
    conn.commit()
    cur.close()
    conn.close()
    return JSONResponse({"ok": True})


@app.post("/materials")
async def add_material_set(request: Request):
    data = await request.json()
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO materials (
          name,
          cr4_1_5, cr4_2_0,
          s275_3_0, s275_4_0, s275_5_0, s275_6_0, s275_8_0, s275_10_0, s275_12_0, s275_15_0, s275_20_0,
          s355_3_0, s355_4_0, s355_5_0, s355_6_0, s355_8_0, s355_10_0, s355_12_0, s355_15_0, s355_20_0,
          hardox450_4_0, hardox450_5_0, hardox450_6_0, hardox450_8_0, hardox450_10_0, hardox450_12_0, hardox450_15_0, hardox450_20_0,
          ss304dp1_0_9, ss304dp1_1_5, ss304dp1_2_0, ss304dp1_3_0, ss304dp1_4_0, ss304dp1_5_0, ss304dp1_6_0, ss304dp1_8_0, ss304dp1_10_0,
          al5251_1_0, al5251_2_0, al5251_3_0, al5251_4_0, al5251_5_0, al5251_6_0, al5251_8_0, al5251_10_0,
          galv_1_0, galv_1_5, galv_2_0, galv_2_5, galv_3_0,
          a1050_1_0, a1050_1_5, a1050_2_0, a1050_3_0, a1050_4_0,
          created_at, is_active
        ) VALUES (
          %s,
          %s,%s,
          %s,%s,%s,%s,%s,%s,%s,%s,%s,
          %s,%s,%s,%s,%s,%s,%s,%s,%s,
          %s,%s,%s,%s,%s,%s,%s,%s,
          %s,%s,%s,%s,%s,%s,%s,%s,%s,
          %s,%s,%s,%s,%s,%s,%s,%s,
          %s,%s,%s,%s,%s,
          %s,%s,%s,%s,%s,
          NOW(), false
        )
    """, (
        data["name"],
        data.get("cr4_1_5"), data.get("cr4_2_0"),
        data.get("s275_3_0"), data.get("s275_4_0"), data.get("s275_5_0"), data.get("s275_6_0"), data.get("s275_8_0"), data.get("s275_10_0"), data.get("s275_12_0"), data.get("s275_15_0"), data.get("s275_20_0"),
        data.get("s355_3_0"), data.get("s355_4_0"), data.get("s355_5_0"), data.get("s355_6_0"), data.get("s355_8_0"), data.get("s355_10_0"), data.get("s355_12_0"), data.get("s355_15_0"), data.get("s355_20_0"),
        data.get("hardox450_4_0"), data.get("hardox450_5_0"), data.get("hardox450_6_0"), data.get("hardox450_8_0"), data.get("hardox450_10_0"), data.get("hardox450_12_0"), data.get("hardox450_15_0"), data.get("hardox450_20_0"),
        data.get("ss304dp1_0_9"), data.get("ss304dp1_1_5"), data.get("ss304dp1_2_0"), data.get("ss304dp1_3_0"), data.get("ss304dp1_4_0"), data.get("ss304dp1_5_0"), data.get("ss304dp1_6_0"), data.get("ss304dp1_8_0"), data.get("ss304dp1_10_0"),
        data.get("al5251_1_0"), data.get("al5251_2_0"), data.get("al5251_3_0"), data.get("al5251_4_0"), data.get("al5251_5_0"), data.get("al5251_6_0"), data.get("al5251_8_0"), data.get("al5251_10_0"),
        data.get("galv_1_0"), data.get("galv_1_5"), data.get("galv_2_0"), data.get("galv_2_5"), data.get("galv_3_0"),
        data.get("a1050_1_0"), data.get("a1050_1_5"), data.get("a1050_2_0"), data.get("a1050_3_0"), data.get("a1050_4_0"),
    ))
    conn.commit()
    cur.close()
    conn.close()
    return JSONResponse({"ok": True})


@app.post("/rates/{rate_id}/select")
async def select_rate(rate_id: int):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("UPDATE burden_rates SET is_active = false")
    cur.execute("UPDATE burden_rates SET is_active = true WHERE id = %s", (rate_id,))
    conn.commit()
    cur.close()
    conn.close()
    return JSONResponse({"ok": True})


@app.post("/rates")
async def add_rate(request: Request):
    data = await request.json()
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        """INSERT INTO burden_rates
           (name, laser_hourly_rate, tube_hourly_rate, fold_rate, weld_saw_rate,
            machine_rate, assembly_rate, finishing_markup, scrap_factor,
            mins_per_fold, fold_setup_mins, created_at, is_active)
           VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s, NOW(), false)""",
        (data["name"], float(data["laser_hourly_rate"]), float(data["tube_hourly_rate"]),
         float(data["fold_rate"]), float(data["weld_saw_rate"]), float(data["machine_rate"]),
         float(data["assembly_rate"]), float(data["finishing_markup"]), float(data["scrap_factor"]),
         float(data["mins_per_fold"]), float(data["fold_setup_mins"]))
    )
    conn.commit()
    cur.close()
    conn.close()
    return JSONResponse({"ok": True})
