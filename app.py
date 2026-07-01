from fastapi import FastAPI, Request, UploadFile, File, Form
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles
from fastapi.responses import JSONResponse, RedirectResponse
from pathlib import Path
from dotenv import load_dotenv
load_dotenv()
import io, os, re, uuid, zipfile
import fitz  # PyMuPDF
from readNestsCSV import parse_nest_csv
from readDrawings import parse_drawing_pdf
import psycopg2
import psycopg2.extras

BASE_DIR = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "uploads"

app = FastAPI()

app.mount("/static",   StaticFiles(directory="static"),        name="static")
app.mount("/uploads",  StaticFiles(directory=str(UPLOAD_DIR)), name="uploads")
templates = Jinja2Templates(directory="templates")
templates.env.filters["stem"] = lambda p: Path(p).stem


def _strip_rev(pn: str) -> str:
    return re.sub(r"-R\d+$", "", pn, flags=re.IGNORECASE).upper()


# CSV material name → active materials table column
_MAT_COL = {
    "SPC1.5": "cr4_1_5",        "SPC2.0": "cr4_2_0",
    "SPH3.0": "s275_3_0",       "SPH4.0": "s275_4_0",
    "SPH5.0": "s275_5_0",       "SPH6.0": "s275_6_0",
    "SPH8.0": "s275_8_0",       "SPH10.0": "s275_10_0",
    "SPH5.0-355-GRADE": "s355_5_0", "SPH6.0-355-GRADE": "s355_6_0",
    "SS400-3.0": "s275_3_0",    "SS400-4.0": "s275_4_0",
    "SS400-5.0": "s275_5_0",    "SS400-6.0": "s275_6_0",
    "SS400-8.0": "s275_8_0",    "SS400-10.0": "s275_10_0",
    "SS400-12.0": "s275_12_0",  "SS400-15.0": "s275_15_0",
    "SS400-20.0": "s275_20_0",
    "ALU-5251-H22-1.5MM": "al5251_1_5", "ALU-5251-H22-2MM": "al5251_2_0",
    "ALU-5251-H22-3MM":   "al5251_3_0", "ALU-5251-H22-5MM": "al5251_5_0",
    "A5052-2.0": "al5251_2_0",  "A5052-3.0": "al5251_3_0",
    "A5052-4.0": "al5251_4_0",  "A5052-5.0": "al5251_5_0",
    "GALV-1.0":  "galv_1_0",    "GALV-1.5":  "galv_1_5",
    "GALV-2.0":  "galv_2_0",    "GALV-2.5":  "galv_2_5",
    "GALV-3.0":  "galv_3_0",
    "A1050-1.0": "a1050_1_0",   "A1050-1.5": "a1050_1_5",
    "A1050-2.0": "a1050_2_0",   "A1050-3.0": "a1050_3_0",
    "A1050-4.0": "a1050_4_0",
}

_MAT_LABEL = {
    "SPC1.5": "Mild steel CR4 1.5mm",      "SPC2.0": "Mild steel CR4 2.0mm",
    "SPH4.0": "Hot rolled S275 4.0mm",     "SPH5.0": "Hot rolled S275 5.0mm",
    "SPH6.0": "Hot rolled S275 6.0mm",     "SPH8.0": "Hot rolled S275 8.0mm",
    "SS400-8.0":  "Hot rolled S275 8.0mm", "SS400-10.0": "Hot rolled S275 10.0mm",
    "SS400-12.0": "Hot rolled S275 12.0mm","SS400-15.0": "Hot rolled S275 15.0mm",
    "SS400-20.0": "Hot rolled S275 20.0mm",
    "SPH5.0-355-GRADE": "Hot rolled S355 5.0mm",
    "SPH6.0-355-GRADE": "Hot rolled S355 6.0mm",
    "ALU-5251-H22-2MM": "Aluminium 5251 H22 2.0mm",
    "ALU-5251-H22-3MM": "Aluminium 5251 H22 3.0mm",
    "ALU-5251-H22-5MM": "Aluminium 5251 H22 5.0mm",
    "GALV-1.5": "Galvanised 1.5mm",        "GALV-2.0": "Galvanised 2.0mm",
    "GALV-2.5": "Galvanised 2.5mm",        "GALV-3.0": "Galvanised 3.0mm",
}


def get_db():
    return psycopg2.connect(os.getenv("DATABASE_STRING"))


def render_pdf_preview(pdf_path: Path, dpi: int = 150) -> Path | None:
    """Render first page of a PDF to a PNG beside it. Returns the PNG path or None."""
    try:
        png_path = pdf_path.with_suffix(".png")
        doc  = fitz.open(str(pdf_path))
        page = doc[0]
        pix  = page.get_pixmap(matrix=fitz.Matrix(dpi / 72, dpi / 72))
        pix.save(str(png_path))
        doc.close()
        return png_path
    except Exception:
        return None


@app.get("/")
async def home(request: Request):
    return templates.TemplateResponse("base.html", {"request": request})


@app.get("/upload")
async def upload(request: Request):
    return templates.TemplateResponse("upload.html", {"request": request})


@app.post("/upload")
async def handle_upload(
    csv_file: UploadFile = File(...),
    zip_file: UploadFile = File(...),
    quote_name: str = Form(...),
    customer: str = Form(default=""),
):
    uid = uuid.uuid4().hex

    # Save CSV
    nest_dir = UPLOAD_DIR / "nests"
    nest_dir.mkdir(parents=True, exist_ok=True)
    csv_path = nest_dir / f"{uid}_{csv_file.filename}"
    csv_path.write_bytes(await csv_file.read())

    # Unzip and save individual PDFs
    drawings_dir = UPLOAD_DIR / "drawings" / uid
    drawings_dir.mkdir(parents=True, exist_ok=True)
    pdf_paths = []
    with zipfile.ZipFile(io.BytesIO(await zip_file.read())) as zf:
        for entry in zf.namelist():
            if entry.lower().endswith(".pdf"):
                safe_name = Path(entry).name
                pdf_path = drawings_dir / safe_name
                pdf_path.write_bytes(zf.read(entry))
                render_pdf_preview(pdf_path)
                pdf_paths.append(pdf_path)

    # Store file paths in DB
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO nests (file_path, imported_at) VALUES (%s, NOW()) RETURNING id",
        (csv_path.relative_to(BASE_DIR).as_posix(),),
    )
    nest_id = cur.fetchone()[0]

    for pdf_path in pdf_paths:
        cur.execute(
            "INSERT INTO drawings (file_path, imported_at) VALUES (%s, NOW())",
            (pdf_path.relative_to(BASE_DIR).as_posix(),),
        )

    conn.commit()
    cur.close()
    conn.close()

    return JSONResponse({"ok": True, "nest_id": nest_id, "pdf_count": len(pdf_paths)})


@app.get("/match/{nest_id}")
async def match_review(request: Request, nest_id: int):
    conn = get_db()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)

    cur.execute("SELECT * FROM nests WHERE id = %s", (nest_id,))
    nest = cur.fetchone()
    if not nest:
        cur.close(); conn.close()
        return JSONResponse({"error": "Not found"}, status_code=404)

    uid = Path(nest["file_path"]).name.split("_")[0]
    cur.execute("SELECT * FROM drawings WHERE file_path LIKE %s",
                (f"uploads/drawings/{uid}/%",))
    batch_drawings = list(cur.fetchall())

    # Parts that already have a quote (pricing history)
    cur.execute("""
        SELECT DISTINCT np.part_number
        FROM nest_parts np JOIN quotes q ON q.nest_part_id = np.id
    """)
    priced = {row["part_number"].upper() for row in cur.fetchall()}
    cur.close(); conn.close()

    # Parse each uploaded PDF → build lookup keyed by stripped part number
    drawing_by_part: dict = {}
    for d in batch_drawings:
        # Filename: "LC_TMC_10_BBRP_1 Blade Bolt Reinf Plate.PDF"
        # Split on first space → part number token + description
        stem = Path(d["file_path"]).stem
        stem_parts = stem.split(" ", 1)
        raw_pn_from_name = stem_parts[0].replace("_", "-")   # underscores → hyphens
        desc_from_name   = stem_parts[1] if len(stem_parts) > 1 else ""

        tb, pf, desc = {}, {}, desc_from_name
        key = _strip_rev(raw_pn_from_name)   # filename-based key as default

        try:
            parsed = parse_drawing_pdf(str(BASE_DIR / d["file_path"]))
            if parsed["pages"]:
                tb   = parsed["pages"][0]["title_block"]
                pf   = parsed["pages"][0]["pricing_factors"]
                desc = desc_from_name or tb.get("description") or ""
                tb_pn = tb.get("part_number", "")
                if tb_pn:
                    key = _strip_rev(tb_pn)  # title block is authoritative when present
        except Exception:
            pass

        drawing_by_part[key] = {"db": d, "tb": tb, "pf": pf, "description": desc}

    # Parse CSV
    csv_result = parse_nest_csv(str(BASE_DIR / nest["file_path"]))

    # Match CSV parts → drawings
    matched_parts = []
    used_ids: set = set()
    for part in csv_result["parts"]:
        key = _strip_rev(part["part_number"])
        match = drawing_by_part.get(key)
        if match:
            used_ids.add(match["db"]["id"])
        matched_parts.append({
            **part,
            "drawing":     match["db"] if match else None,
            "description": match["description"] if match else "",
            "has_pricing": key in priced or part["part_number"].upper() in priced,
        })

    unmatched_drawings = [
        {"id": d["id"], "name": Path(d["file_path"]).stem}
        for d in batch_drawings if d["id"] not in used_ids
    ]

    stats = {
        "total":            len(matched_parts),
        "drawings_matched": sum(1 for p in matched_parts if p["drawing"]),
        "no_drawing":       sum(1 for p in matched_parts if not p["drawing"]),
        "not_priced":       sum(1 for p in matched_parts if not p["has_pricing"]),
    }

    return templates.TemplateResponse("match.html", {
        "request":            request,
        "nest_id":            nest_id,
        "parts":              matched_parts,
        "unmatched_drawings": unmatched_drawings,
        "stats":              stats,
        "csv_warnings":       csv_result.get("_warnings", []),
    })


@app.post("/match/{nest_id}/confirm")
async def confirm_match(request: Request, nest_id: int):
    form = await request.form()

    conn = get_db()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    cur.execute("SELECT * FROM nests WHERE id = %s", (nest_id,))
    nest = cur.fetchone()

    csv_result = parse_nest_csv(str(BASE_DIR / nest["file_path"]))

    wc = conn.cursor()

    # Update nests row with CSV metadata
    wc.execute("""
        UPDATE nests SET source_type = 'csv', schedule_name = %s, total_process_qty = %s
        WHERE id = %s
    """, (csv_result["source_file"], csv_result["schedule"]["total_process_qty"], nest_id))

    for part in csv_result["parts"]:
        pn = part["part_number"]
        drawing_id_str = form.get(f"drawing_{pn}", "")
        drawing_id = int(drawing_id_str) if drawing_id_str and drawing_id_str.strip() else None

        # Update drawings table with parsed PDF metadata
        if drawing_id:
            cur.execute("SELECT file_path FROM drawings WHERE id = %s", (drawing_id,))
            d_row = cur.fetchone()
            if d_row:
                try:
                    parsed = parse_drawing_pdf(str(BASE_DIR / d_row["file_path"]))
                    if parsed["pages"]:
                        tb = parsed["pages"][0]["title_block"]
                        pf = parsed["pages"][0]["pricing_factors"]
                        notes = ", ".join(pf.get("process_notes", []) or []) or None
                        wc.execute("""
                            UPDATE drawings SET
                                part_number = %s, revision = %s, description = %s,
                                material = %s, thickness_mm = %s, finish = %s,
                                status = %s, drawn_by = %s,
                                bending_required = %s, bend_count = %s, process_notes = %s
                            WHERE id = %s
                        """, (
                            tb.get("part_number"), tb.get("revision"), tb.get("description"),
                            tb.get("material"), tb.get("thickness_mm"), tb.get("finish"),
                            tb.get("status"), tb.get("drawn_by"),
                            pf.get("bending_required"), pf.get("bend_count"), notes,
                            drawing_id,
                        ))
                except Exception:
                    pass

        wc.execute("""
            INSERT INTO nest_parts (
                nest_id, drawing_id, part_number,
                size_x_mm, size_y_mm, thickness_mm, material,
                process_time_seconds, total_qty, order_qty
            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """, (
            nest_id, drawing_id, pn,
            part.get("size_x_mm"), part.get("size_y_mm"), part.get("thickness_mm"),
            part.get("material"), part.get("process_time_per_part_seconds"),
            part.get("total_qty"), part.get("total_qty"),
        ))

    conn.commit()
    cur.close(); wc.close(); conn.close()

    return RedirectResponse(url=f"/price/{nest_id}", status_code=303)


# ── Pricing pages ────────────────────────────────────────────────────────────

@app.get("/price/{nest_id}")
async def price_start(nest_id: int):
    return RedirectResponse(url=f"/price/{nest_id}/1", status_code=302)


@app.get("/price/{nest_id}/review")
async def price_review(request: Request, nest_id: int):
    conn = get_db()
    cur  = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    cur.execute("""
        SELECT np.id, np.part_number, np.material, np.total_qty,
               d.description AS drawing_desc
        FROM nest_parts np
        LEFT JOIN drawings d ON d.id = np.drawing_id
        WHERE np.nest_id = %s ORDER BY np.id
    """, (nest_id,))
    parts = [dict(p) for p in cur.fetchall()]
    cur.close(); conn.close()
    return templates.TemplateResponse("review.html", {
        "request": request, "nest_id": nest_id, "parts": parts,
    })


@app.post("/price/{nest_id}/save")
async def price_save(request: Request, nest_id: int):
    data = await request.json()
    conn = get_db()
    cur  = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    cur.execute("SELECT id FROM burden_rates WHERE is_active = true LIMIT 1")
    rates_row = cur.fetchone()
    cur.execute("SELECT id FROM materials WHERE is_active = true LIMIT 1")
    mat_row   = cur.fetchone()
    wc = conn.cursor()
    for p in data:
        np_id  = p.get("nest_part_id")
        if not np_id:
            continue
        active = p.get("active_processes", [])
        wc.execute("""
            INSERT INTO quotes (
                nest_part_id, burden_rate_id, material_id, quoted_at, quantity,
                purchased_total, finish_cost_per_part, sticker_cost, delivery_cost,
                misc_cost, margin_pct,
                tube_active, weld_active, saw_active, machine_active, assembly_active,
                tube_cut_time_mins, tube_kg_per_metre, tube_length_m, tube_cost_per_kg,
                weld_time_mins, saw_num_cuts, saw_mins_per_cut, saw_setup_mins,
                machine_time_mins, assembly_time_mins
            ) VALUES (
                %s,%s,%s,NOW(),%s, %s,%s,%s,%s,%s,%s,
                %s,%s,%s,%s,%s,
                %s,%s,%s,%s, %s, %s,%s,%s, %s,%s
            )
        """, (
            np_id,
            rates_row["id"] if rates_row else None,
            mat_row["id"]   if mat_row   else None,
            p.get("qty"),
            p.get("purchased", 0), p.get("finish_cost", 0),
            p.get("sticker_cost", 0.07), p.get("delivery_cost", 5),
            p.get("misc_cost", 0),       p.get("margin_pct", 10),
            "tube" in active, "weld" in active, "saw"      in active,
            "machine" in active, "assembly" in active,
            p.get("tube_cut_time_mins"), p.get("tube_kg_per_metre"),
            p.get("tube_length_m"),      p.get("tube_cost_per_kg"),
            p.get("weld_time_mins"),
            p.get("saw_num_cuts"),       p.get("saw_mins_per_cut"),
            p.get("saw_setup_mins"),     p.get("machine_time_mins"),
            p.get("assembly_time_mins"),
        ))
    conn.commit()
    cur.close(); wc.close(); conn.close()
    return JSONResponse({"ok": True})


@app.post("/price/calculate")
async def price_calculate(request: Request):
    d    = await request.json()
    conn = get_db()
    cur  = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    cur.execute("SELECT * FROM burden_rates WHERE is_active = true LIMIT 1")
    rates = cur.fetchone()
    cur.close(); conn.close()
    if not rates:
        return JSONResponse({"error": "No active burden rates"}, status_code=400)

    def f(k, default=0):
        try: return float(d.get(k) or default)
        except: return float(default)
    def iv(k, default=0):
        try: return int(d.get(k) or default)
        except: return int(default)

    qty    = max(iv("qty", 1), 1)
    active = d.get("active_processes", [])

    laser  = ((d.get("size_x_mm", 0) * d.get("size_y_mm", 0) / 1_000_000)
               * f("material_cost_m2") * rates["scrap_factor"] * qty
               + (rates["laser_hourly_rate"] / 60) * (f("process_time_seconds") / 60) * qty)

    fold   = 0.0
    nf     = iv("num_folds")
    if nf > 0:
        mf   = f("mins_per_fold", 0.25)
        fs   = f("fold_setup_mins", 5)
        fold = (qty * nf * mf + nf * mf + fs) * (rates["fold_rate"] / 60)

    tube = (f("tube_cut_time_mins") * (rates["tube_hourly_rate"] / 60) * qty
            + f("tube_kg_per_metre") * f("tube_length_m") * f("tube_cost_per_kg")
            * rates["scrap_factor"] * qty) if "tube" in active else 0.0

    weld     = (f("weld_time_mins") / 60) * rates["weld_saw_rate"] * qty if "weld" in active else 0.0
    saw      = ((f("saw_setup_mins") + iv("saw_num_cuts") * f("saw_mins_per_cut"))
                * (rates["weld_saw_rate"] / 60)) if "saw" in active else 0.0
    machine  = (f("machine_time_mins")  / 60) * rates["machine_rate"]  * qty if "machine"  in active else 0.0
    assembly = (f("assembly_time_mins") / 60) * rates["assembly_rate"] * qty if "assembly" in active else 0.0

    finish   = f("finish_cost") * rates["finishing_markup"] * qty
    purchase = f("purchased")   * qty
    sticker  = f("sticker_cost",  0.07) * qty
    delivery = f("delivery_cost", 5.0)  * qty
    misc     = f("misc_cost")           * qty

    total        = laser + fold + tube + weld + saw + machine + assembly + finish + purchase + sticker + delivery + misc
    margin_pct   = f("margin_pct", 10)
    cost_per_part = round(total / qty, 2)
    margin_total  = round(total * (1 + margin_pct / 100), 2)

    return JSONResponse({
        "cost_per_part": cost_per_part,
        "total_cost":    round(total, 2),
        "margin_total":  margin_total,
        "breakdown": {
            "laser": round(laser, 2), "fold":     round(fold, 2),
            "tube":  round(tube, 2),  "weld":     round(weld, 2),
            "saw":   round(saw, 2),   "machine":  round(machine, 2),
            "assembly": round(assembly, 2),
        },
    })


@app.get("/price/{nest_id}/{part_index}")
async def price_part(request: Request, nest_id: int, part_index: int):
    conn = get_db()
    cur  = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    cur.execute("""
        SELECT np.*, d.file_path AS drawing_path, d.description AS drawing_desc,
               d.revision AS drawing_revision, d.bending_required, d.bend_count
        FROM nest_parts np
        LEFT JOIN drawings d ON d.id = np.drawing_id
        WHERE np.nest_id = %s ORDER BY np.id
    """, (nest_id,))
    all_parts = [dict(p) for p in cur.fetchall()]

    if not all_parts or part_index < 1 or part_index > len(all_parts):
        cur.close(); conn.close()
        return JSONResponse({"error": "Not found"}, status_code=404)

    part = all_parts[part_index - 1]

    col          = _MAT_COL.get(part["material"])
    material_cost = None
    if col:
        cur.execute(f"SELECT {col} FROM materials WHERE is_active = true LIMIT 1")
        row = cur.fetchone()
        if row:
            material_cost = row[col]

    cur.close(); conn.close()

    drawing_image_url = None
    if part.get("drawing_path"):
        png_rel = part["drawing_path"].rsplit(".", 1)[0] + ".png"
        if (BASE_DIR / png_rel).exists():
            drawing_image_url = "/" + png_rel

    return templates.TemplateResponse("price.html", {
        "request":           request,
        "nest_id":           nest_id,
        "part_index":        part_index,
        "total_parts":       len(all_parts),
        "part":              part,
        "part_numbers":      [p["part_number"] for p in all_parts],
        "material_cost":     material_cost,
        "catalogue":         _MAT_LABEL.get(part["material"], ""),
        "drawing_image_url": drawing_image_url,
    })


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
