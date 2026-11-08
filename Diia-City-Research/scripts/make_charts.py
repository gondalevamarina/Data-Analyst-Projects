"""Побудова графіків 1-5 з data/diia_city_for_tableau.xlsx. Запуск з кореня проєкту: python scripts/make_charts.py"""
import pandas as pd, numpy as np, matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.ticker import PercentFormatter

F="data/diia_city_for_tableau.xlsx"
c=pd.read_excel(F,"Компанії",dtype={"ЄДРПОУ":str}); m=pd.read_excel(F,"Помісячно"); co=pd.read_excel(F,"Когорти")
OUT="charts/"
BLUE,RED,GRAY,GREEN="#1F4E78","#C0504D","#8C8C8C","#4F8A5B"
plt.rcParams.update({"font.family":"DejaVu Sans","font.size":10,"axes.spines.top":False,"axes.spines.right":False,"axes.grid":True,"grid.alpha":.25,"axes.axisbelow":True})
SRC="Джерело: реєстр Дія.City (city.diia.gov.ua), збір 08.11.2026. Листопад 2026 — неповний місяць (дані до 05.11)."
def fin(fig,ax,title,sub,name):
    fig.suptitle(title,x=.01,ha="left",fontsize=14,fontweight="bold")
    ax.set_title(sub,loc="left",fontsize=10,color="#555")
    fig.text(.01,.01,SRC,fontsize=7.5,color="#777")
    fig.tight_layout(rect=(0,.03,1,.95)); fig.savefig(OUT+name,dpi=170); plt.close(fig)
def monthax(ax):
    ax.xaxis.set_major_locator(mdates.MonthLocator(bymonth=[1,7])); ax.xaxis.set_major_formatter(mdates.DateFormatter("%m.%Y"))

last=m["Місяць"].max()
# 1 нові резиденти
fig,ax=plt.subplots(figsize=(11,5.2))
cols=[BLUE if d!=last else "#9DB7D1" for d in m["Місяць"]]
ax.bar(m["Місяць"],m["Нові резиденти"],width=24,color=cols); monthax(ax)
ax.set_ylabel("Нових резидентів за місяць")
yr=c.groupby("Рік набуття").size()
ax.text(.01,.95,"Усього за рік: "+"  |  ".join(f"{y}: {v}" for y,v in yr.items()),transform=ax.transAxes,fontsize=9,va="top")
fin(fig,ax,"1. Зростання Дія.City: нові резиденти за місяцями","Кількість компаній, що набули статусу резидента","chart1_new_residents.png")

# 2 втрати
fig,ax=plt.subplots(figsize=(11,5.2))
cols=[RED if d!=last else "#E3B3B1" for d in m["Місяць"]]
ax.bar(m["Місяць"],m["Втратили статус"],width=24,color=cols); monthax(ax)
ax.set_ylabel("Втратили статус за місяць")
q=c.dropna(subset=["Дата втрати"]).groupby(c["Дата втрати"].dt.to_period("Q")).size()
ax.text(.01,.95,f"Найбільше за квартал: {q.idxmax()} — {q.max()} компаній",transform=ax.transAxes,fontsize=9,va="top")
fin(fig,ax,"2. Припинення резидентства: втрати статусу за місяцями","Дата втрати = дата останнього рішення Мінцифри для компаній зі статусом «Не резидент»","chart2_lost_status.png")

# 3 структура: стартапи за роком набуття
g=c.groupby("Рік набуття")["Стартап"].value_counts().unstack(fill_value=0)
sh=g.div(g.sum(axis=1),axis=0)
fig,ax=plt.subplots(figsize=(11,5.2))
x=np.arange(len(g)); ax.bar(x,sh["Так"],color=BLUE,label="Стартап"); ax.bar(x,sh["Ні"],bottom=sh["Так"],color="#BFCFE0",label="Не стартап")
for i,(yy,row) in enumerate(g.iterrows()):
    ax.text(i,sh["Так"].iloc[i]/2,f"{sh['Так'].iloc[i]:.0%}\n({row['Так']})",ha="center",va="center",color="white",fontsize=9)
    ax.text(i,sh["Так"].iloc[i]+sh["Ні"].iloc[i]/2,f"{sh['Ні'].iloc[i]:.0%}\n({row['Ні']})",ha="center",va="center",fontsize=9)
ax.set_xticks(x); ax.set_xticklabels(g.index); ax.yaxis.set_major_formatter(PercentFormatter(1)); ax.set_xlabel("Рік набуття статусу"); ax.legend(loc="upper center",ncol=2,bbox_to_anchor=(.5,-.12),frameon=False)
fin(fig,ax,"3. Хто входить до Дія.City: частка стартапів за роком набуття статусу","У дужках кількість компаній. Форма ТОВ має 99,4% рядків, тому за формами розбивку не показано","chart3_startups.png")

# 4 утримання
fig,ax=plt.subplots(figsize=(11,5.2))
w=.26; x=np.arange(len(co))
for k,(n,col) in enumerate([(1,BLUE),(2,GREEN),(3,"#E0A030")]):
    v=co[f"Частка {n} р."]
    b=ax.bar(x+(k-1)*w,v.fillna(0),w,color=col,label=f"Через {n} {'рік' if n==1 else 'роки'}")
    for i,val in enumerate(v):
        if pd.notna(val): ax.text(x[i]+(k-1)*w,val+.01,f"{val:.0%}",ha="center",fontsize=8.5)
ax.set_xticks(x); ax.set_xticklabels([f"{y}\n(n={n})" for y,n in zip(co["Когорта (рік набуття)"],co["Розмір когорти"])])
ax.set_ylim(0,1.08); ax.yaxis.set_major_formatter(PercentFormatter(1)); ax.legend(loc="upper right",frameon=False)
ax.set_xlabel("Когорта: рік набуття статусу (n — розмір когорти)")
fin(fig,ax,"4. Стійкість резидентів: частка, що зберегла статус","Лише для компаній, у яких минуло достатньо часу; порожньо — спостереження ще надто коротке","chart4_retention.png")

# 5 головні показники
fig,(ax,ax2)=plt.subplots(2,1,figsize=(11,7),gridspec_kw={"height_ratios":[1.6,1]},sharex=True)
ax.plot(m["Місяць"],m["Резидентів на кінець місяця"],color=BLUE,lw=2.2)
ax.fill_between(m["Місяць"],m["Резидентів на кінець місяця"],color=BLUE,alpha=.12)
ax.set_ylabel("Чинних резидентів"); ax.annotate(f"{m['Резидентів на кінець місяця'].iloc[-1]:,}".replace(","," "),(m["Місяць"].iloc[-1],m["Резидентів на кінець місяця"].iloc[-1]),xytext=(-50,-4),textcoords="offset points",fontweight="bold")
ax2.bar(m["Місяць"],m["Чиста зміна"],width=24,color=[GREEN if v>=0 else RED for v in m["Чиста зміна"]]); ax2.set_ylabel("Чиста зміна за місяць"); monthax(ax2)
fin(fig,ax,"5. Головні показники: кількість резидентів і чиста зміна","Чинні резиденти на кінець місяця (угорі) та різниця «нові мінус втрати» (унизу)","chart5_overview.png")

