"""
generador_stock.py
==================
Procesa los archivos Excel y genera el HTML con datos inyectados.

Uso:
    python generador_stock.py

    (o con parámetros explícitos)
    python generador_stock.py --movimientos Reporte.xlsx --stock "PosicionStock Table.xls"

Requiere:
    pip install pandas openpyxl xlrd
"""

import argparse
import json
import re
import subprocess
import sys
import tempfile
import zipfile
from datetime import datetime, date
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    sys.exit("Instalar pandas: pip install pandas openpyxl")


# ── Configuración ──────────────────────────────────────────────────────────────

DOC_VENTAS = "Remito de Venta Planta"

PRODUCTOS = [
    "Advance Gato Almohaditas x 8kg","Advance Gatito y Gata Lactante Kitten All Breed x 7.5 kg",
    "Advance Perro Cachorro All Breed x 1.5 kg","Advance Perro Almohaditas x 15kg",
    "Advance Gato Adulto Indoor x 8 kg","Avena Pelada x 25 Kgs","Guardian Perro Adulto x 15 kg",
    "Colza x 20 Kgs","Fandog x 20 kg","Advance Bio Cookies Pollo y Arroz X 150 Gr",
    "Guardian Cachorro x 10 kg","Arroz p/Perro x 20 kg","Nutricare Mordida Grande x 7.5 kg",
    "Guardian Perro Adulto x 20 kg","Advance Perro Adulto Weight Control x 15 kg",
    "Guardian Gato Adulto x 10 kg","Advance Gato Adulto Salmon All Breed x 8 kg",
    "Miau Gato Adulto x 8kg","Don Dog x 20 kg","Alpiste x 25 Kgs",
    "Advance Hipoalergénico Large Breed (Cordero y Arroz) x 15 kg",
    "Advance Gato Adulto All Breed Urinary x 8kg",
    "Advance Hipoalergénico Small Breed (Cordero y Arroz) x 12 kg",
    "Nutricare Mordida Grande x 20 kg","Advance Perro Cachorro All Breed x 12 kg",
    "Ultramix Cachorros x 10 kg","Advance Perro Adulto Large Breed x 20 kg",
    "Mezcla Gallina x 25 Kgs","Sustituto +Lait Plus x 25 Kg","Girasol Confitero x 20 Kg",
    "Advance Perro Adulto Small Breed x 12 kg","Mezcla Semillero x 25 Kgs",
    "Avena Pelada x 20 Kgs","Nutricare Cachorro x 15 kg","Mezcla Canario x 25 Kgs",
    "Ultramix Gato x 10 kg","Vitamina x 25 Kg","Guardian Perro Adulto Raza Pequeña x 15 kg",
    "Nutricare Mordida Pequeña x 7.5 kg","Piedras Sanitarias X 25 Kg",
    "Polenta para perro x 20 Kgs","Nutricare Mordida Pequeña x 20 kg",
    "Ultramix Adulto x 15 kg","Advance Bio Cookies Carne y Arroz X 150 Gr",
    "Advance Bio Cookies Cordero y Arroz X 150 Gr","Arroz p/Perro Saborizado x 25 kg",
    "Girasol Rayado x 30 Kg","Mezcla Jaulon x 25 Kgs","Mezcla Pájaro Chico x 25 Kg",
    "Mijo x 25 Kgs","Sorgo x 25 Kgs",
    # HFM
    "HFM Alpiste x 20 kg","HFM Arroz sab carne y veg x 15 kg","HFM Avena pelada x 25 kg",
    "HFM Girasol chico x 20 kg","HFM Girasol confitero x 20 kg",
    "HFM Mezcla canarios especial con vitaminas x 20 kg","HFM Mezcla de gallinas x 20 kg",
    "HFM Mezcla mil granos especial x 20 kg","HFM Mezcla pajaro especial con vitaminas x 20 kg",
    "HFM Mijo x 20 kg","HFM Sorgo x 30 kg","HFM Vitaminas x 20 kg",
]

ABC_MAP = {
    "Guardian Gato Adulto x 10 kg":"A","Advance Perro Adulto Large Breed x 20 kg":"A",
    "Mezcla Gallina x 25 Kgs":"A","Advance Perro Adulto Small Breed x 12 kg":"A",
    "Nutricare Mordida Grande x 20 kg":"A","Piedras Sanitarias X 25 Kg":"A",
    "Nutricare Mordida Pequeña x 20 kg":"A","Sustituto +Lait Plus x 25 Kg":"A",
    "Advance Perro Cachorro All Breed x 12 kg":"A",
    "Guardian Perro Adulto x 20 kg":"B","Advance Gato Adulto Salmon All Breed x 8 kg":"B",
    "Advance Hipoalergénico Large Breed (Cordero y Arroz) x 15 kg":"B",
    "Advance Gato Adulto All Breed Urinary x 8kg":"B","Girasol Confitero x 20 Kg":"B",
    "Mezcla Semillero x 25 Kgs":"B","Nutricare Cachorro x 15 kg":"B",
    "Mezcla Canario x 25 Kgs":"B","Ultramix Gato x 10 kg":"A",
    "Guardian Perro Adulto Raza Pequeña x 15 kg":"B","Nutricare Mordida Pequeña x 7.5 kg":"B",
    "Girasol Rayado x 30 Kg":"B","Mezcla Jaulon x 25 Kgs":"B",
    "Mezcla Pájaro Chico x 25 Kg":"B","Mijo x 25 Kgs":"B",
}


# ── Lectura de XLS vía LibreOffice ─────────────────────────────────────────────

def read_excel(path: str) -> pd.DataFrame:
    """Lee .xlsx, .xls o .csv automáticamente. Repara xlsx malformados."""
    import warnings
    p = Path(path)

    if p.suffix.lower() == '.csv':
        # Detectar separador (coma o punto y coma)
        with open(path, 'r', encoding='utf-8-sig') as f:
            primera = f.readline()
        sep = ';' if primera.count(';') > primera.count(',') else ','
        return pd.read_csv(path, encoding='utf-8-sig', sep=sep)

    if p.suffix.lower() == '.xls':
        return pd.read_excel(path, engine='xlrd')

    # .xlsx
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        try:
            return pd.read_excel(path, engine='openpyxl')
        except TypeError:
            # Reparar typo 'biltinId' que generan algunos sistemas de gestión
            print(f"  (reparando {p.name}...)")
            fixed = Path(tempfile.gettempdir()) / f"_fixed_{p.name}"
            with zipfile.ZipFile(path, 'r') as zin:
                with zipfile.ZipFile(fixed, 'w', zipfile.ZIP_DEFLATED) as zout:
                    for item in zin.namelist():
                        data = zin.read(item)
                        if item == 'xl/styles.xml':
                            data = data.replace(b'biltinId', b'builtinId')
                        zout.writestr(item, data)
            return pd.read_excel(fixed, engine='openpyxl')


# ── Proceso de datos ───────────────────────────────────────────────────────────

def find_col(df: pd.DataFrame, candidates: list) -> str | None:
    cols_lower = {c.strip().lower(): c for c in df.columns}
    for c in candidates:
        if c.lower() in cols_lower:
            return cols_lower[c.lower()]
    return None


def fmt_date(d) -> str | None:
    if pd.isna(d):
        return None
    if isinstance(d, (datetime, date)):
        return d.strftime('%d/%m/%Y')
    return str(d)


def diff_days(d) -> int | None:
    if pd.isna(d):
        return None
    if isinstance(d, datetime):
        d = d.date()
    if isinstance(d, date):
        return (d - date.today()).days
    return None


def detectar_quiebre(serie_meses: list, mes_actual_idx: int, stock_actual: float) -> bool:
    """
    Detecta posible quiebre de stock con dos señales:
    1. Stock actual en cero (evidencia directa de quiebre vigente)
    2. Caída brusca de ventas entre dos meses consecutivos dentro de los
       últimos 4 meses completos (ej: vendía 20/mes y de repente cayó a 2,
       sin necesidad de que se haya "recuperado" después)
    serie_meses: lista de 12 dicts {label, qty, qtyPrev} indexada 0=Ene .. 11=Dic
    mes_actual_idx: índice (0-11) del mes en curso, para no evaluarlo (puede estar incompleto)
    stock_actual: stock actual del producto
    """
    if stock_actual == 0:
        return True

    inicio = max(0, mes_actual_idx - 4)
    ventana = [serie_meses[i]['qty'] for i in range(inicio, mes_actual_idx)]
    if len(ventana) < 2:
        return False

    # Caída brusca entre meses consecutivos: el mes anterior tenía venta
    # relevante y el siguiente cayó a menos del 25% de ese valor
    for i in range(len(ventana) - 1):
        antes, despues = ventana[i], ventana[i + 1]
        if antes >= 3 and despues <= antes * 0.25:
            return True

    return False


def procesar(mov_path: str, stock_path: str) -> dict:
    print(f"  Leyendo movimientos: {mov_path}")
    df_mov = read_excel(mov_path)

    print(f"  Leyendo stock: {stock_path}")
    df_stock = read_excel(stock_path)

    CF = find_col(df_mov, ['Fecha', 'fecha'])
    CP = find_col(df_mov, ['Producto', 'producto'])
    CD = find_col(df_mov, ['Documento nombre', 'Documento'])
    CC = find_col(df_mov, ['Cantidad principal', 'Cantidad'])
    CPS = find_col(df_stock, ['Producto actual', 'Producto', 'producto'])
    CVT = find_col(df_stock, ['Vto. (Lote)', 'Vencimiento', 'Vto'])
    CCS = find_col(df_stock, ['Cantidad', 'cantidad'])

    assert CF and CP and CD and CC, f"Columnas faltantes en movimientos. Disponibles: {list(df_mov.columns)}"
    assert CPS and CCS, f"Columnas faltantes en stock. Disponibles: {list(df_stock.columns)}"

    df_mov[CF] = pd.to_datetime(df_mov[CF], errors='coerce')
    ventas = df_mov[
        (df_mov[CD].astype(str).str.strip() == DOC_VENTAS) &
        (pd.to_numeric(df_mov[CC], errors='coerce') < 0)
    ].copy()
    assert len(ventas) > 0, "Sin ventas en el reporte."

    max_f = ventas[CF].max()
    print(f"  Fecha máxima de ventas: {max_f.date()}")

    t30 = max_f - pd.Timedelta(days=30)
    t90 = max_f - pd.Timedelta(days=90)
    tA0 = max_f - pd.Timedelta(days=395)
    tA1 = max_f - pd.Timedelta(days=365)

    ventas['_cant'] = pd.to_numeric(ventas[CC], errors='coerce').abs()
    ventas['_prod'] = ventas[CP].astype(str).str.strip()

    def agg(mask):
        return ventas[mask].groupby('_prod')['_cant'].sum().to_dict()

    u30 = agg(ventas[CF] > t30)
    u90 = agg(ventas[CF] > t90)
    a30 = agg((ventas[CF] > tA0) & (ventas[CF] <= tA1))

    # Gráfico dual: enero–diciembre del año actual + mismo mes año anterior
    MESES_ES = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic']
    ano_actual = max_f.year
    ano_prev = ano_actual - 1

    # Sumar ventas por (año, mes)
    ventas['_ym'] = ventas[CF].dt.year * 100 + ventas[CF].dt.month
    por_mes = ventas.groupby(['_prod', '_ym'])['_cant'].sum()

    ventas_mensuales: dict[str, list] = {}
    for prod in ventas['_prod'].unique():
        serie = []
        tiene_datos = False
        for m in range(1, 13):
            qty = float(por_mes.get((prod, ano_actual * 100 + m), 0))
            qty_prev = float(por_mes.get((prod, ano_prev * 100 + m), 0))
            if qty or qty_prev:
                tiene_datos = True
            serie.append({
                'label': MESES_ES[m - 1],
                'qty': round(qty, 1),
                'qtyPrev': round(qty_prev, 1),
            })
        if tiene_datos:
            ventas_mensuales[prod] = serie

    # Stock
    stock_map: dict[str, float] = {}
    lotes_map: dict[str, dict] = {}
    for _, row in df_stock.iterrows():
        prod = str(row.get(CPS, '') or '').strip()
        if not prod:
            continue
        cant = float(row.get(CCS, 0) or 0)
        if cant <= 0:
            continue
        stock_map[prod] = stock_map.get(prod, 0) + cant
        vto_str = fmt_date(row.get(CVT)) if CVT else None
        vto_dias = diff_days(row.get(CVT)) if CVT else None
        if vto_str:
            if prod not in lotes_map:
                lotes_map[prod] = {}
            if vto_str not in lotes_map[prod]:
                lotes_map[prod][vto_str] = {'vto': vto_str, 'cant': 0, 'dias': vto_dias}
            lotes_map[prod][vto_str]['cant'] += cant

    # Armar rows
    mes_actual_idx = max_f.month - 1  # 0-indexed
    rows = []
    for prod in PRODUCTOS:
        stock = round(stock_map.get(prod, 0))
        s30 = round(u30.get(prod, 0) * 10) / 10
        s90 = round(u90.get(prod, 0) * 10) / 10
        sa30 = round(a30.get(prod, 0) * 10) / 10
        prom = round(s90 / 3 * 10) / 10
        abc = ABC_MAP.get(prod, 'C')
        lotes = sorted(lotes_map.get(prod, {}).values(), key=lambda x: x['dias'] if x['dias'] is not None else 9999)
        lotes = [{'vto': l['vto'], 'cant': round(l['cant']), 'dias': l['dias']} for l in lotes]
        tend = round((s30 - prom) / prom * 100) if prom > 0 else None
        serie = ventas_mensuales.get(prod)
        quiebre = detectar_quiebre(serie, mes_actual_idx, stock) if serie else (stock == 0)
        rows.append({
            'prod': prod, 'stock': stock, 'abc': abc,
            's30': s30, 'prom': prom, 'sa30': sa30,
            'tend': tend, 'lotes': lotes, 'quiebre': quiebre,
        })

    return {
        'rows': rows,
        'ventasMensuales': ventas_mensuales,
        'fechaStr': fmt_date(max_f),
        'productos': PRODUCTOS,
        'mode': 'generated',
    }


# ── Inyectar en HTML ───────────────────────────────────────────────────────────

def inyectar(html: str, data: dict) -> str:
    json_str = json.dumps(data, ensure_ascii=False)
    new_line = f'const INJECTED_DATA = {json_str};'
    return re.sub(r'const INJECTED_DATA = null;.*', new_line, html)


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description='Genera HTML con datos inyectados')
    parser.add_argument('--movimientos', default='Reporte.xlsx')
    parser.add_argument('--stock', default='PosicionStock_Table.xlsx')
    parser.add_argument('--template', default='StockReventa_v11.html')
    parser.add_argument('--output', default='index.html')
    args = parser.parse_args()

    template = Path(args.template)
    assert template.exists(), f"Template no encontrado: {template}"

    mov = Path(args.movimientos)
    assert mov.exists(), f"Archivo no encontrado: {mov}"

    stk = Path(args.stock)
    assert stk.exists(), f"Archivo no encontrado: {stk}"

    print("Procesando datos...")
    data = procesar(str(mov), str(stk))
    print(f"  {len(data['rows'])} productos procesados")

    html = template.read_text(encoding='utf-8')
    html = inyectar(html, data)
    Path(args.output).write_text(html, encoding='utf-8')
    print(f"✅ HTML generado: {args.output}")


if __name__ == '__main__':
    main()
