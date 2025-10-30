# ==========================================================
# PROJECT: PENJADWALAN & PENGGAJIAN KARYAWAN
# ==========================================================
# Tujuan: Optimasi jadwal kerja & biaya tenaga kerja
# Tools: Python (PuLP) vs Excel Solver
# ==========================================================

import pandas as pd
from pulp import LpProblem, LpVariable, LpMinimize, lpSum, LpBinary, value

# ----------------------------------------------------------
# 1. DATA KARYAWAN (20 SAMPEL)
# ----------------------------------------------------------
data = {
    "ID": [f"K{i:02d}" for i in range(1, 21)],
    "Nama": [
        "Andi", "Budi", "Citra", "Dewa", "Eka", "Fina", "Gani", "Hana", "Irma", "Joko",
        "Kiki", "Lala", "Made", "Nia", "Oki", "Putra", "Qori", "Rina", "Sari", "Tono"
    ],
    "Gaji_per_Jam": [
        40000, 35000, 45000, 30000, 32000, 38000, 30000, 35000, 33000, 37000,
        30000, 36000, 32000, 38000, 33000, 34000, 30000, 36000, 40000, 42000
    ],
    "Maks_Jam_Mingguan": [
        40, 40, 40, 32, 32, 40, 32, 40, 32, 40, 32, 40, 32, 40, 32, 40, 32, 40, 40, 40
    ],
}

df = pd.DataFrame(data)

# ----------------------------------------------------------
# 2. PARAMETER PENJADWALAN
# ----------------------------------------------------------
shifts = ["Pagi", "Siang", "Malam"]
jam_per_shift = 8

# Kebutuhan minimal karyawan per shift
kebutuhan_shift = {
    "Pagi": 5,
    "Siang": 5,
    "Malam": 3
}

# ----------------------------------------------------------
# 3. MODEL OPTIMASI DENGAN PULP
# ----------------------------------------------------------
model = LpProblem("Optimasi_Penjadwalan", LpMinimize)

# Variabel biner: 1 jika karyawan i bekerja di shift j
x = LpVariable.dicts("x", [(i, j) for i in df.index for j in shifts], 0, 1, LpBinary)

# Fungsi tujuan: minimisasi total biaya
model += lpSum(df.loc[i, "Gaji_per_Jam"] * jam_per_shift * x[(i, j)]
               for i in df.index for j in shifts)

# Kendala: kebutuhan minimal per shift
for j in shifts:
    model += lpSum(x[(i, j)] for i in df.index) >= kebutuhan_shift[j]

# Kendala: setiap karyawan hanya 1 shift per hari
for i in df.index:
    model += lpSum(x[(i, j)] for j in shifts) <= 1

# Kendala: total jam ≤ batas maksimum
for i in df.index:
    model += lpSum(x[(i, j)] * jam_per_shift for j in shifts) <= df.loc[i, "Maks_Jam_Mingguan"]

# ----------------------------------------------------------
# 4. SOLUSI MODEL
# ----------------------------------------------------------
model.solve()

# ----------------------------------------------------------
# 5. HASIL OPTIMASI PYTHON
# ----------------------------------------------------------
hasil_opt = []
for i in df.index:
    for j in shifts:
        if x[(i, j)].value() == 1:
            biaya = df.loc[i, "Gaji_per_Jam"] * jam_per_shift
            hasil_opt.append([df.loc[i, "Nama"], j, df.loc[i, "Gaji_per_Jam"], jam_per_shift, biaya])

hasil_df = pd.DataFrame(hasil_opt, columns=["Nama", "Shift", "Gaji/Jam", "Jam", "Biaya"])
total_biaya_python = hasil_df["Biaya"].sum()

# ----------------------------------------------------------
# 6. SIMULASI HASIL EXCEL SOLVER (Manual Formula)
# ----------------------------------------------------------
# Asumsikan Excel memilih pekerja dengan biaya terendah untuk memenuhi kebutuhan minimum
excel_simulasi = df.nsmallest(sum(kebutuhan_shift.values()), "Gaji_per_Jam")
excel_simulasi["Shift"] = ["Pagi"]*5 + ["Siang"]*5 + ["Malam"]*3 + ["Cadangan"]*0
excel_simulasi["Biaya"] = excel_simulasi["Gaji_per_Jam"] * jam_per_shift
total_biaya_excel = excel_simulasi["Biaya"].sum()

# ----------------------------------------------------------
# 7. PERBANDINGAN HASIL
# ----------------------------------------------------------
perbandingan = pd.DataFrame({
    "Metode": ["Python (PuLP)", "Excel (Solver)"],
    "Total Biaya (Rp)": [total_biaya_python, total_biaya_excel]
})

# Tambahkan efisiensi (% penghematan)
efisiensi = (1 - total_biaya_python / total_biaya_excel) * 100
perbandingan["Efisiensi (%)"] = [round(efisiensi, 2), 0]

# ----------------------------------------------------------
# 8. SIMPAN KE EXCEL
# ----------------------------------------------------------
with pd.ExcelWriter("perbandingan_penjadwalan.xlsx") as writer:
    df.to_excel(writer, sheet_name="Data Karyawan", index=False)
    hasil_df.to_excel(writer, sheet_name="Optimasi_Python", index=False)
    excel_simulasi.to_excel(writer, sheet_name="Simulasi_Excel", index=False)
    perbandingan.to_excel(writer, sheet_name="Perbandingan", index=False)

# ----------------------------------------------------------
# 9. OUTPUT KE TERMINAL
# ----------------------------------------------------------
print("\n=== HASIL OPTIMASI PYTHON (PuLP) ===")
print(hasil_df)
print("\nTotal Biaya Minimum (Python): Rp", f"{int(total_biaya_python):,}".replace(",", "."))

print("\n=== HASIL SIMULASI EXCEL ===")
print(excel_simulasi[["Nama", "Shift", "Gaji_per_Jam", "Biaya"]])
print("\nTotal Biaya Excel: Rp", f"{int(total_biaya_excel):,}".replace(",", "."))

print("\n=== PERBANDINGAN EFISIENSI ===")
print(perbandingan)
print("\nFile Excel disimpan sebagai: perbandingan_penjadwalan.xlsx")