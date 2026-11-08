"""Підготовка даних: очищення, помісячна динаміка, когорти. Запуск з кореня проєкту: python scripts/prepare_data.py"""
import pandas as pd, numpy as np
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

SRC="data/diia_city_registry_raw.xlsx"
OUT="data/diia_city_for_tableau.xlsx"
ASOF=pd.Timestamp("2026-11-08")

r=pd.read_excel(SRC,"Реєстр",dtype={"ЄДРПОУ":str})
d=pd.read_excel(SRC,"Рішення",dtype={"ЄДРПОУ":str})

# --- очищення
opf_map={"товариство з обмеженою відповідальністю":"Товариство з обмеженою відповідальністю",
"приватне підприємство":"Приватне підприємство","акціонерне товариство":"Акціонерне товариство",
"дочірне підпримство":"Дочірнє підприємство","приватне акціонерне товариство":"Приватне акціонерне товариство",
"науково-дослідне підприємство":"Науково-дослідне підприємство"}
r["ОПФ"]=r["Орг.-правова форма"].str.strip().str.lower().map(opf_map).fillna(r["Орг.-правова форма"])
r["Назва (очищена)"]=r["Назва компанії"].str.strip().str.strip('"“”«»\'').str.strip()

r["Чинний"]=(r["Статус на сайті"]=="Резидент").astype(int)
r["Дата набуття"]=r["Дата набуття статусу"]
r["Дата втрати"]=r["Дата втрати статусу"]
r["Рік набуття"]=r["Дата набуття"].dt.year
r["Місяць набуття"]=r["Дата набуття"].dt.to_period("M").dt.to_timestamp()
r["Рік втрати"]=r["Дата втрати"].dt.year
r["Місяць втрати"]=r["Дата втрати"].dt.to_period("M").dt.to_timestamp()
end=r["Дата втрати"].fillna(ASOF)
r["Тривалість, днів"]=(end-r["Дата набуття"]).dt.days
r["Повторний вхід"]=np.where(r["ЄДРПОУ"].duplicated(keep=False),"Так","Ні")
r["Кількість рішень"]=r["Кількість рішень"].fillna(0).astype(int)

comp=r[["№ з/п","ЄДРПОУ","Назва (очищена)","Назва компанії","ОПФ","Статус на сайті","Чинний","Стартап",
"Відповідність доходу","Дата набуття","Місяць набуття","Рік набуття","Дата втрати","Місяць втрати",
"Рік втрати","Тривалість, днів","Повторний вхід","Кількість рішень"]]

# --- події (довга таблиця)
base=r[["ЄДРПОУ","Назва (очищена)","ОПФ","Стартап","Повторний вхід"]]
a=base.assign(Дата=r["Дата набуття"],Подія="Набуття",Зміна=1)
l=base[r["Дата втрати"].notna()].assign(Дата=r.loc[r["Дата втрати"].notna(),"Дата втрати"],Подія="Втрата",Зміна=-1)
ev=pd.concat([a,l]).sort_values(["Дата","Подія"]).reset_index(drop=True)
ev["Місяць"]=ev["Дата"].dt.to_period("M").dt.to_timestamp(); ev["Рік"]=ev["Дата"].dt.year
ev=ev[["Дата","Місяць","Рік","Подія","Зміна","ЄДРПОУ","Назва (очищена)","ОПФ","Стартап","Повторний вхід"]]

# --- помісячна динаміка
months=pd.period_range(ev["Місяць"].min(),ev["Місяць"].max(),freq="M").to_timestamp()
m=pd.DataFrame({"Місяць":months})
m["Нові резиденти"]=m["Місяць"].map(ev[ev.Подія=="Набуття"].groupby("Місяць").size()).fillna(0).astype(int)
m["Втратили статус"]=m["Місяць"].map(ev[ev.Подія=="Втрата"].groupby("Місяць").size()).fillna(0).astype(int)
m["Чиста зміна"]=m["Нові резиденти"]-m["Втратили статус"]
m["Резидентів на кінець місяця"]=m["Чиста зміна"].cumsum()
assert m["Резидентів на кінець місяця"].iloc[-1]==r["Чинний"].sum(), "cumsum != resident count"

# --- когорти утримання
rows=[]
for y,g in r.groupby("Рік набуття"):
    row={"Когорта (рік набуття)":y,"Розмір когорти":len(g)}
    for n in (1,2,3):
        elig=g[g["Дата набуття"]<=ASOF-pd.Timedelta(days=365*n)]
        kept=elig[(elig["Дата втрати"].isna())|((elig["Дата втрати"]-elig["Дата набуття"]).dt.days>=365*n)]
        row[f"Придатних для {n} р."]=len(elig); row[f"Утримали {n} р."]=len(kept)
        row[f"Частка {n} р."]=(len(kept)/len(elig)) if len(elig) else None
    rows.append(row)
coh=pd.DataFrame(rows)

info=pd.DataFrame({"Примітки":[
"Джерело: city.diia.gov.ua/registry/resident, збір 08.11.2026. Підготовлено з diia_city_registry_full.xlsx.",
f"Станом на {ASOF:%d.%m.%Y}: рядків (епізодів резидентства) {len(r)}, чинних {int(r['Чинний'].sum())}, втратили статус {int((1-r['Чинний']).sum())}.",
"«Дата втрати» = дата останнього рішення Мінцифри для компаній зі статусом «Не резидент». Вибірково звірте з PDF-рішеннями.",
"Повторний вхід = компанія з тим самим ЄДРПОУ має кілька рядків (втратила статус і набула його знову). Кожен рядок = окремий епізод.",
"«Тривалість, днів»: для чинних рахується до дати станом на, для втрачених до дати втрати.",
"Когорти: утримання за N років рахується лише для компаній, що набули статус щонайменше N років тому (стовпець «Придатних»).",
"У когорті 2025-2026 років спостерігати 2-3 роки ще неможливо, тому там частка порожня.",
"Аркуші: Компанії (1 рядок = 1 епізод), Події (набуття/втрата, для помісячних графіків), Помісячно, Когорти, Рішення."]})

def write(w,df,name,datecols=(),pct=()):
    df.to_excel(w,sheet_name=name,index=False)
    ws=w.sheets[name]
    for c in ws[1]:
        c.font=Font(name="Arial",size=10,bold=True,color="FFFFFF")
        c.fill=PatternFill("solid",fgColor="1F4E78"); c.alignment=Alignment(wrap_text=True,vertical="center")
    for i,col in enumerate(df.columns,1):
        ws.column_dimensions[get_column_letter(i)].width=min(max(12,len(str(col))+4),46)
        for cell in ws[get_column_letter(i)][1:]:
            cell.font=Font(name="Arial",size=10)
            if col in datecols: cell.number_format="DD.MM.YYYY"
            if col in pct: cell.number_format="0.0%"
            if col=="ЄДРПОУ": cell.number_format="@"
    ws.freeze_panes="A2"; ws.auto_filter.ref=ws.dimensions

with pd.ExcelWriter(OUT,engine="openpyxl") as w:
    write(w,comp,"Компанії",datecols=("Дата набуття","Місяць набуття","Дата втрати","Місяць втрати"))
    write(w,ev,"Події",datecols=("Дата","Місяць"))
    write(w,m,"Помісячно",datecols=("Місяць",))
    write(w,coh,"Когорти",pct=("Частка 1 р.","Частка 2 р.","Частка 3 р."))
    d2=d.copy(); d2["Дата рішення"]=pd.to_datetime(d2["Дата рішення"])
    write(w,d2,"Рішення",datecols=("Дата рішення",))
    info.to_excel(w,sheet_name="Про файл",index=False); w.sheets["Про файл"].column_dimensions["A"].width=120
    for c in w.sheets["Про файл"]["A"]: c.font=Font(name="Arial",size=10)

