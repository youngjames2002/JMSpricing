from fastapi import FastAPI, Request
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles
from dotenv import load_dotenv
load_dotenv()
import os
import psycopg2
from jinja2 import Environment, FileSystemLoader

app = FastAPI()

app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

@app.get("/")
async def home(request: Request):
    return templates.TemplateResponse("base.html", {"request": request})

@app.get("/rates")
async def rates(request: Request):
    conn = psycopg2.connect(os.getenv("DATABASE_STRING"))
    cur = conn.cursor()
    cur.execute("SELECT * FROM materials;")
    mat_rows = cur.fetchall()
    cur.execute("SELECT * FROM burden_rates;")
    rates_rows = cur.fetchall()
    cur.close()
    return templates.TemplateResponse("index.html", {"request": request, "materials":mat_rows, "burden_rates": rates_rows})
