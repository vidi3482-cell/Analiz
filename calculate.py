# -*- coding: utf-8 -*-
# Расчёт: сколько уроков в неделю ставить до конца 1 полугодия.
# Запуск: python calculate.py data reports
import sys, os, re, math, glob, datetime
import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

def week_weight(code):
    if code is None: return 0.0
    s = str(code).strip().upper().replace("Ё", "Е")
    if s in ("О", "Т"): return 1.0
    parts = re.split(r"[/\\]", s)
    if len(parts) > 1:
        return 0.5 if any(p.strip() in ("О", "Т") for p in parts) else 0.0
    return 0.0  # К, Э, С, Пп, УП, Уп, ТП, А...

def fw(x):
    f, c = math.floor(x + 1e-9), math.ceil(x - 1e-9)
    return str(c) if f >= c else str(f) + "-" + str(c)

def main(src="data", outdir="reports"):
    os.makedirs(outdir, exist_ok=True)
    nfiles = sorted(glob.glob(os.path.join(src, "Нагрузка_*.xlsx")))
    pfile = os.path.join(src, "Plan.xlsx")
    if not nfiles or not os.path.exists(pfile):
        raise SystemExit("Нужны Нагрузка_ГГГГ-ММ-ДД.xlsx и Plan.xlsx в папке " + src)
    nfile = nfiles[-1]
    m = re.search(r"(\d{4}-\d{2}-\d{2})", os.path.basename(nfile))
    rd = datetime.date.fromisoformat(m.group(1)) if m else datetime.date.today()

    df = pd.read_excel(nfile, sheet_name="По группам и предметам")
    df.columns = ["teacher","group","subject","tarif","fact","rest","plan1",
                  "fact1","plan2","fact2","theory","sub1","sub2","status"]
    df = df[df.teacher.notna() & df.group.notna() & df.subject.notna()]
    df = df[df.teacher != "Преподаватель"]
    for c in ["tarif","plan1","fact1"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df = df[df.plan1.notna()].copy()
    df["group"] = df["group"].astype(str).str.strip()
    df["need"] = (df.plan1 - df.fact1).astype(int)

    wb = openpyxl.load_workbook(pfile, data_only=True)
    cal, curric = {}, {}
    for g in wb.sheetnames:
        ws = wb[g]
        cal[g] = [ws.cell(2, c).value for c in range(2, 54)]
        for r in range(4, ws.max_row + 1):
            subj, hrs = ws.cell(r, 1).value, ws.cell(r, 2).value
            if subj and isinstance(hrs, (int, float)):
                curric[(g, str(subj).strip())] = hrs

    sep1 = datetime.date(rd.year if rd.month >= 9 else rd.year - 1, 9, 1)
    week1 = sep1 - datetime.timedelta(days=sep1.weekday())
    boundary = datetime.date(rd.year, 12, 28)
    cur_week = (rd - week1).days // 7 + 1
    last_week = (boundary - week1).days // 7
    weeks_idx = list(range(cur_week + 1, last_week + 1))
    print("Отчёт:", rd, "| граница:", boundary, "| остались недели:", weeks_idx)

    avail = {g: sum(week_weight(codes[i-1]) for i in weeks_idx if 1 <= i <= 52)
             for g, codes in cal.items()}
    default_weeks = float(len(weeks_idx))
    df["weeks"] = df["group"].map(avail).fillna(default_weeks)
    df["pw"] = df["need"].clip(lower=0) / df["weeks"].replace(0, float("nan"))

    par = {}
    for (g, s), sg in df.groupby(["group", "subject"]):
        p = False
        if len(sg) > 1:
            gh = curric.get((g, s)); tot = sg.tarif.sum()
            if gh and tot > gh * 1.25: p = True
            elif sg[["sub1","sub2"]].notna().any().any() and s in (
                "Иностранный язык", "Физическая культура и здоровье"): p = True
        for i in sg.index: par[i] = p
    df["par"] = pd.Series(par)

    rows = []
    for g, gdf in df.groupby("group"):
        act = gdf[gdf.need > 0]; slots = 0.0
        for s, sg in act.groupby("subject"):
            slots += sg.need.max() if (sg.par.any() and len(sg) > 1) else sg.need.sum()
        if slots == 0: continue
        w = avail.get(g, default_weeks)
        rows.append({"group": g, "weeks": w, "lessons": int(act.need.sum()),
                     "slots": int(slots),
                     "spw": round(slots/w, 1) if w else None,
                     "lpd": round(slots/(w*5), 1) if w else None})
    sdf = pd.DataFrame(rows).sort_values("group")

    out = Workbook()
    hdr = PatternFill("solid", fgColor="1F4E78")
    red = PatternFill("solid", fgColor="FFC7CE")
    orange = PatternFill("solid", fgColor="FFE0B2")
    yellow = PatternFill("solid", fgColor="FFF2CC")
    green = PatternFill("solid", fgColor="C6EFCE")
    def style_header(ws, row=1):
        for c in ws[row]:
            if c.value is not None:
                c.font = Font(bold=True, color="FFFFFF"); c.fill = hdr
                c.alignment = Alignment(wrap_text=True, vertical="center")

    ws = out.active; ws.title = "Сводка по группам"
    ws.append(["Группа","Недель осталось","Осталось уроков","Слотов",
               "Слотов/нед","Уроков/день","Оценка"])
    for _, r in sdf.iterrows():
        d = r.lpd if r.lpd is not None else 99
        v = ("РЕАЛЬНО при 5-6/день" if d <= 6 else
             "НА ПРЕДЕЛЕ 6-7/день" if d <= 7 else "нужен перенос части плана")
        ws.append([r.group, r.weeks, r.lessons, r.slots, r.spw, r.lpd, v])
    style_header(ws)
    for row in ws.iter_rows(min_row=2):
        d = row[5].value or 99
        row[6].fill = green if d <= 6 else (yellow if d <= 7 else red)
    for i, w in enumerate([9,13,14,10,11,11,32], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"

    ws2 = out.create_sheet("По предметам")
    ws2.append(["Группа","Предмет","Преподаватель","План 1п","Факт 1п",
                "Остаток","Недель","Уроков/нед точно","Ставить/нед","Параллельно"])
    for _, r in df.sort_values(["group","need","subject"],
                               ascending=[True,False,True]).iterrows():
        ws2.append([r.group, r.subject, r.teacher, int(r.plan1), int(r.fact1),
                    int(r.need), r.weeks,
                    round(r.pw, 1) if r.need > 0 else 0,
                    fw(r.pw) if r.need > 0 else "0",
                    "да" if r.par else ""])
    style_header(ws2)
    for row in ws2.iter_rows(min_row=2):
        need, pw = row[5].value, row[7].value
        if need > 0:
            row[8].fill = red if pw > 10 else (orange if pw > 6 else yellow)
        else:
            row[8].fill = green
    for i, w in enumerate([9,48,20,8,8,9,8,12,12,11], 1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    ws2.freeze_panes = "D2"; ws2.auto_filter.ref = ws2.dimensions

    for _, s in sdf.iterrows():
        g = s.group
        gdf = df[df.group == g].sort_values(["need","subject"], ascending=[False,True])
        wg = out.create_sheet(g)
        wg.append([g + ": осталось " + str(s.weeks) + " нед | " + str(s.slots)
                   + " слотов ~ " + str(s.spw) + "/нед (" + str(s.lpd) + "/день)"])
        wg.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
        wg.cell(1,1).font = Font(bold=True, size=12)
        wg.append(["Предмет","Преподаватель","Остаток","Ставить/нед","Параллельно","Статус"])
        style_header(wg, 2)
        for _, r in gdf.iterrows():
            if r.plan1 == 0 and r.fact1 == 0: continue
            st = ("нужно добрать" if r.need > 0 else
                  "выполнено" if r.need == 0 else "аванс")
            wg.append([r.subject, r.teacher, int(r.need),
                       fw(r.pw) if r.need > 0 else "0",
                       "да" if r.par else "", st])
        for row in wg.iter_rows(min_row=3):
            need = row[2].value
            if isinstance(need, int) and need > 0:
                try: v = float(str(row[3].value).split("-")[-1])
                except ValueError: v = 0
                row[3].fill = red if v > 10 else (orange if v > 6 else yellow)
            elif isinstance(need, int):
                row[3].fill = green
        for i, w in enumerate([48,20,9,12,12,16], 1):
            wg.column_dimensions[get_column_letter(i)].width = w
        wg.freeze_panes = "A3"

    fname = os.path.join(outdir, "Сколько_уроков_в_неделю_до_"
                         + boundary.strftime("%d.%m") + ".xlsx")
    out.save(fname)
    sdf.to_csv(os.path.join(outdir, "svodka.csv"), sep=";", index=False,
               encoding="utf-8-sig", decimal=",")
    cols = df[df.need > 0][["group","subject","teacher","need","weeks","pw","par"]].copy()
    cols["pw"] = cols["pw"].round(1)
    cols.to_csv(os.path.join(outdir, "po_predmetam.csv"), sep=";", index=False,
                encoding="utf-8-sig", decimal=",")
    print("Готово:", fname)
    print(sdf.to_string(index=False))

if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "data",
         sys.argv[2] if len(sys.argv) > 2 else "reports")