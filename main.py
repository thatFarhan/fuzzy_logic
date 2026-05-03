import openpyxl

def read_file(path_file):
    wb = openpyxl.load_workbook(path_file)
    worksheet = wb.active

    data = []
    header = [cell.value for cell in worksheet[1]]

    for row in worksheet.iter_rows(min_row=2, values_only=True):
        record = dict(zip(header, row))

        if record.get('id Pelanggan') is not None:
            data.append({
                'id'       : int(record['id Pelanggan']),
                'pelayanan': float(record['Pelayanan']),
                'harga'    : float(record['harga'])
            })
    return data

def left_trapezoid(x, a, b, c, d):
    if x <= b:
        return 1.0
    elif b < x < d:
        return (d - x) / (d - b)
    else:
        return 0.0

def triangle(x, a, b, c):
    if x <= a or x >= c:
        return 0.0
    elif a < x <= b:
        return (x - a) / (b - a)
    else:
        return (c - x) / (c - b)

def right_trapezoid(x, a, b, c, d):
    if x <= a:
        return 0.0
    elif a < x < b:
        return (x - a) / (b - a)
    else:
        return 1.0

def fuzzify_service(p):
    return {
        'buruk' : left_trapezoid(p, 1, 1, 20, 40),
        'sedang': triangle(p, 20, 50, 80),
        'baik'  : right_trapezoid(p, 60, 80, 100, 100)
    }

def fuzzify_price(h):
    return {
        'murah' : left_trapezoid(h, 20000, 20000, 28000, 40000),
        'sedang': triangle(h, 28000, 37500, 47000),
        'mahal' : right_trapezoid(h, 40000, 47000, 55000, 55000)
    }

def inference(mu_service, mu_price):
    p = mu_service
    h = mu_price

    r1 = min(p['baik'],   h['murah'])   # Sangat_Layak
    r2 = min(p['baik'],   h['sedang'])  # Layak
    r3 = min(p['baik'],   h['mahal'])   # Cukup_Layak
    r4 = min(p['sedang'], h['murah'])   # Layak
    r5 = min(p['sedang'], h['sedang'])  # Cukup_Layak
    r6 = min(p['sedang'], h['mahal'])   # Tidak_Layak
    r7 = min(p['buruk'],  h['murah'])   # Cukup_Layak
    r8 = min(p['buruk'],  h['sedang'])  # Tidak_Layak
    r9 = min(p['buruk'],  h['mahal'])   # Tidak_Layak

    return {
        'tidak_layak' : max(r6, r8, r9),
        'cukup_layak' : max(r3, r5, r7),
        'layak'       : max(r2, r4),
        'sangat_layak': r1
    }

def mu_output_tidak_layak(z):
    return left_trapezoid(z, 0, 0, 15, 35)

def mu_output_cukup_layak(z):
    return triangle(z, 20, 40, 60)

def mu_output_layak(z):
    return triangle(z, 45, 65, 85)

def mu_output_sangat_layak(z):
    return right_trapezoid(z, 65, 85, 100, 100)

def defuzzifikasi(alpha, n_titik=1000):
    z_min, z_max = 0.0, 100.0
    step = (z_max - z_min) / n_titik

    pembilang = 0.0
    penyebut  = 0.0

    z = z_min
    while z <= z_max:
        mu_tl = min(alpha['tidak_layak'],  mu_output_tidak_layak(z))
        mu_cl = min(alpha['cukup_layak'],  mu_output_cukup_layak(z))
        mu_l  = min(alpha['layak'],        mu_output_layak(z))
        mu_sl = min(alpha['sangat_layak'], mu_output_sangat_layak(z))

        mu_agg = max(mu_tl, mu_cl, mu_l, mu_sl)

        pembilang += z * mu_agg
        penyebut  += mu_agg
        z += step

    if penyebut == 0:
        return 0.0
    return pembilang / penyebut

def hitung_skor(restoran):
    mu_p = fuzzify_service(restoran['pelayanan'])
    mu_h = fuzzify_price(restoran['harga'])
    alpha = inference(mu_p, mu_h)
    skor  = defuzzifikasi(alpha)
    return round(skor, 4)

def process_all(data):
    hasil = []
    for r in data:
        skor = hitung_skor(r)
        hasil.append({
            'id'       : r['id'],
            'pelayanan': r['pelayanan'],
            'harga'    : r['harga'],
            'skor'     : skor
        })
    hasil.sort(key=lambda x: x['skor'], reverse=True)
    return hasil[:5]

def save_output(top5, path_output):
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "Peringkat Restoran"

    headers = ["Peringkat", "ID Restoran", "Kualitas Pelayanan (1-100)",
               "Harga (Rp)", "Skor Kelayakan (0-100)"]
    worksheet.append(headers)

    for rank, r in enumerate(top5, 1):
        worksheet.append([
            rank,
            r['id'],
            r['pelayanan'],
            r['harga'],
            r['skor']
        ])
 
        worksheet.cell(row=rank + 1, column=4).number_format = '#,##0'
        worksheet.cell(row=rank + 1, column=5).number_format = '0.0000'
 
    workbook.save(path_output)
    print(f"Output disimpan ke: {path_output}")

if __name__ == "__main__":
    INPUT_FILE  = "restoran.xlsx"
    OUTPUT_FILE = "top5_restoran.xlsx"

    data = read_file(INPUT_FILE)
    top5 = process_all(data)

    print("5 Restoran Terbaik:\n")
    print(f"  {'Peringkat':<10} {'ID':>4}  {'Pelayanan':>10}  {'Harga (Rp)':>12}  {'Skor':>8}")
    print("  " + "-" * 52)
    for rank, r in enumerate(top5, 1):
        print(f"  {rank:<10} {r['id']:>4}  {r['pelayanan']:>10.0f}  {r['harga']:>12,.0f}  {r['skor']:>8.4f}")

    save_output(top5, OUTPUT_FILE)