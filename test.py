import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np

# === Настройка стиля графиков ===
sns.set_theme(style="whitegrid")
plt.rcParams.update({
    'font.size': 10,
    'axes.titlesize': 14,
    'axes.labelsize': 12,
    'figure.figsize': (12, 8)
})

# === Загрузка данных ===
file_path = "тест номера.xlsx"
df = pd.read_excel(file_path, sheet_name="Выгрузка")

# Проверка наличия нужных колонок
required_cols = ["Куратор", "Наличие номера в лиде", "Сверка клиентских номеров"]
if not all(col in df.columns for col in required_cols):
    raise ValueError(f"Не найдены обязательные колонки: {required_cols}")

# Ожидаемые категории
expected_cats = ["Номер клиента", "Номер агента", "Номер отсутствует"]

# === Функция построения сводной таблицы ===
def build_summary(df, group_col, value_col, table_name):
    pivot = pd.crosstab(
        df[group_col],
        df[value_col],
        margins=True,
        margins_name="Общий итог"
    )
    
    # Добавляем недостающие категории
    for cat in expected_cats:
        if cat not in pivot.columns:
            pivot[cat] = 0
    
    # Порядок столбцов
    pivot = pivot[expected_cats + ["Общий итог"]]
    
    # Считаем % KPI
    pivot["% KPI"] = (pivot["Номер клиента"] / pivot["Общий итог"] * 100).round(2)
    pivot["% KPI"] = pivot["% KPI"].fillna(0)
    
    print(f"\n{'='*60}")
    print(f"{table_name}")
    print(f"{'='*60}")
    print(pivot.to_string())
    
    return pivot

# === Строим сводные ===
summary1 = build_summary(
    df,
    group_col="Куратор",
    value_col="Наличие номера в лиде",
    table_name="% Наличие клиентских номеров в лиде"
)

summary2 = build_summary(
    df,
    group_col="Куратор",
    value_col="Сверка клиентских номеров",
    table_name="% Сверка с партнерской БД"
)

# === Сохраняем в Excel ===
with pd.ExcelWriter("сводные_таблицы_по_номерам.xlsx", engine="openpyxl") as writer:
    summary1.to_excel(writer, sheet_name="Наличие номера в лиде")
    summary2.to_excel(writer, sheet_name="Сверка с партнерской БД")

print("\n✅ Сводные таблицы сохранены в файл: 'сводные_таблицы_по_номерам.xlsx'")

# === График 1: % Наличие клиентских номеров в лиде ===
summary1_plot = summary1.drop("Общий итог").sort_values("% KPI", ascending=False)

plt.figure(figsize=(12, 10))
bars = plt.barh(summary1_plot.index, summary1_plot["% KPI"], color=sns.color_palette("Blues", len(summary1_plot)))
plt.title('% Наличие клиентских номеров в лиде по кураторам', fontsize=16, fontweight='bold')
plt.xlabel('% KPI (доля "Номер клиента")')
plt.ylabel('Куратор')
plt.grid(axis='x', linestyle='--', alpha=0.7)

# Подписи на барах
for bar in bars:
    width = bar.get_width()
    plt.text(width + 0.5, bar.get_y() + bar.get_height()/2, f'{width:.1f}%', 
             va='center', fontsize=9, color='black')

plt.tight_layout()
plt.savefig('график_наличие_номеров.png', dpi=300, bbox_inches='tight')
plt.show()

# === График 2: % Сверка с партнерской БД ===
summary2_plot = summary2.drop("Общий итог").sort_values("% KPI", ascending=False)

plt.figure(figsize=(12, 10))
bars = plt.barh(summary2_plot.index, summary2_plot["% KPI"], color=sns.color_palette("Greens", len(summary2_plot)))
plt.title('% Сверка с партнерской БД по кураторам', fontsize=16, fontweight='bold')
plt.xlabel('% KPI (доля "Номер клиента")')
plt.ylabel('Куратор')
plt.grid(axis='x', linestyle='--', alpha=0.7)

for bar in bars:
    width = bar.get_width()
    plt.text(width + 0.5, bar.get_y() + bar.get_height()/2, f'{width:.1f}%', 
             va='center', fontsize=9, color='black')

plt.tight_layout()
plt.savefig('график_сверка_с_бд.png', dpi=300, bbox_inches='tight')
plt.show()

# === Donut-чарты (опционально) ===
def plot_donut(data, title, filename):
    total = data.loc["Общий итог", expected_cats]
    colors = sns.color_palette("Set2", len(expected_cats))
    
    plt.figure(figsize=(8, 8))
    wedges, texts, autotexts = plt.pie(
        total,
        labels=expected_cats,
        autopct='%1.1f%%',
        startangle=90,
        colors=colors,
        pctdistance=0.85,
        textprops={'fontsize': 11}
    )
    
    # Делаем текст процентов белым и жирным
    for autotext in autotexts:
        autotext.set_color('white')
        autotext.set_fontweight('bold')
    
    # Donut-эффект
    centre_circle = plt.Circle((0, 0), 0.70, fc='white')
    fig = plt.gcf()
    fig.gca().add_artist(centre_circle)
    
    plt.title(title, fontsize=16, fontweight='bold', pad=20)
    plt.axis('equal')
    plt.tight_layout()
    plt.savefig(filename, dpi=300, bbox_inches='tight')
    plt.show()

# Строим donut-чарты
plot_donut(summary1, "Распределение номеров: Наличие в лиде", "donut_наличие.png")
plot_donut(summary2, "Распределение номеров: Сверка с БД", "donut_сверка.png")

print("\n✅ Все графики сохранены в PNG-файлы.")