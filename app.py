import streamlit as st
import pandas as pd
import folium
from streamlit_folium import st_folium
from folium.plugins import Search
import requests  # untuk OSRM

# =========================
# Page Settings
# =========================
st.set_page_config(page_title="Dashboard Visit KC Grand Wisata", layout="wide")
st.title("📍 Dashboard Visit Bank Mandiri KC Grand Wisata")

# =========================
# Load Data from Google Sheets
# =========================
SHEET_ID = "1Z5Ff44r5T9oCR9McTGUg07jmlb3wfjXjFrPj_aMM4SU"
csv_url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"

df = pd.read_csv(csv_url)

# Hapus kolom 'Column 12' jika ada
if "Column 12" in df.columns:
    df = df.drop(columns=["Column 12"])

# =========================
# Clean & Combine Product Data
# =========================

# Gabungkan produk untuk nasabah duplikat
combined_products = (
    df.groupby(["Status Nasabah", "Nama Nasabah / PIC Usaha", "Nama Usaha"])["Produk yang Ditawarkan"]
      .apply(lambda x: ", ".join(sorted(set(x.dropna()))))
      .reset_index()
)

# Merge kembali hasil penggabungan
# Pastikan kolom "Produk yang Ditawarkan" ada sebelum di-drop
if "Produk yang Ditawarkan" in df.columns:
    df = df.drop(columns=["Produk yang Ditawarkan"]).merge(
        combined_products,
        on=["Status Nasabah", "Nama Nasabah / PIC Usaha", "Nama Usaha"],
        how="left"
    )
else:
    # fallback: merge using combined_products into df (avoid error)
    df = df.merge(
        combined_products,
        on=["Status Nasabah", "Nama Nasabah / PIC Usaha", "Nama Usaha"],
        how="left"
    )

# Hapus duplikasi penuh setelah penggabungan (jika kolom ada)
drop_subset = [
    "Status Nasabah",
    "Nama Nasabah / PIC Usaha",
    "Nama Usaha",
    "Produk yang Ditawarkan"
]
existing_subset = [c for c in drop_subset if c in df.columns]
if existing_subset:
    df = df.drop_duplicates(subset=existing_subset, keep="first")

# =========================
# Split Coordinates
# =========================
coord_col = "Koordinat (Latitude, Longitude), Contoh (-6.287825727208808, 107.0433026643262)"

# Pastikan kolom koordinat ada
if coord_col in df.columns:
    # Hanya ambil baris yang mengandung koma dan tidak kosong
    df = df[df[coord_col].notna() & df[coord_col].str.contains(",")]

    # Split koordinat
    df[["Latitude", "Longitude"]] = df[coord_col].str.split(",", expand=True)

    # Convert to float + drop yang gagal konversi
    df["Latitude"] = pd.to_numeric(df["Latitude"], errors="coerce")
    df["Longitude"] = pd.to_numeric(df["Longitude"], errors="coerce")

    # Drop baris yang koordinatnya NaN
    df = df.dropna(subset=["Latitude", "Longitude"])
else:
    st.error(f"Kolom koordinat tidak ditemukan: '{coord_col}'")
    st.stop()

# =========================
# Sidebar Filters
# =========================
st.sidebar.header("Filter Data")

pegawai_filter = st.sidebar.multiselect(
    "Nama Pegawai:",
    df["Nama Pegawai"].dropna().unique()
)

status_nasabah_filter = st.sidebar.multiselect(
    "Status Nasabah:",
    df["Status Nasabah"].dropna().unique()
)

status_closing_filter = st.sidebar.multiselect(
    "Status Closing:",
    df["Status Closing"].dropna().unique()
)

# Produk dropdown based on combined Produk yang Ditawarkan (safe guard)
if "Produk yang Ditawarkan" in df.columns:
    produk_options = sorted(set(", ".join(df["Produk yang Ditawarkan"].dropna()).split(", ")))
else:
    produk_options = []

produk_filter = st.sidebar.multiselect(
    "Produk:",
    produk_options
)

# =========================
# Apply Filters
# =========================
filtered_df = df.copy()

if pegawai_filter:
    filtered_df = filtered_df[filtered_df["Nama Pegawai"].isin(pegawai_filter)]

if status_nasabah_filter:
    filtered_df = filtered_df[filtered_df["Status Nasabah"].isin(status_nasabah_filter)]

if status_closing_filter:
    filtered_df = filtered_df[filtered_df["Status Closing"].isin(status_closing_filter)]

if produk_filter and "Produk yang Ditawarkan" in filtered_df.columns:
    filtered_df = filtered_df[
        filtered_df["Produk yang Ditawarkan"].str.contains("|".join(produk_filter), case=False, na=False)
    ]

# =========================
# Rute: Sidebar selector (harus setelah filtered_df dibuat)
# =========================
st.sidebar.subheader("🧭 Rute Visit")
list_titik = list(filtered_df["Nama Usaha"].dropna().unique())
start_point = st.sidebar.selectbox("Titik Awal:", ["-"] + list_titik)
end_point = st.sidebar.selectbox("Titik Tujuan:", ["-"] + list_titik)

# =========================
# Fungsi panggil OSRM
# =========================
def get_route_osrm(lat1, lon1, lat2, lon2):
    """
    Memanggil OSRM public server untuk rute driving.
    Mengembalikan list [lat, lon] points atau None jika gagal.
    """
    url = f"http://router.project-osrm.org/route/v1/driving/{lon1},{lat1};{lon2},{lat2}?overview=full&geometries=geojson"
    try:
        r = requests.get(url, timeout=8)
        r.raise_for_status()
        data = r.json()
        coords = data["routes"][0]["geometry"]["coordinates"]
        # OSRM returns [lon, lat] -> convert ke [lat, lon]
        return [[c[1], c[0]] for c in coords]
    except Exception as e:
        # debug optional: st.write(e)
        return None

# =========================
# Map Section
# =========================
st.subheader("🗺 Peta Visit Nasabah")

# Base Map (centered sekitar Grand Wisata)
m = folium.Map(location=[-6.287825727208808, 107.0433026643262], zoom_start=14)

# Marker group for Search feature
marker_group = folium.FeatureGroup(name="Visit Points")
m.add_child(marker_group)

# Add Markers
for _, row in filtered_df.iterrows():
    popup = f"""
    <b>Nama Nasabah:</b> {row['Nama Nasabah / PIC Usaha']}<br>
    <b>Nama Usaha:</b> {row['Nama Usaha']}<br>
    <b>Pegawai:</b> {row['Nama Pegawai']}<br>
    <b>Status Nasabah:</b> {row['Status Nasabah']}<br>
    <b>Status Closing:</b> {row['Status Closing']}<br>
    <b>Produk:</b> {row.get('Produk yang Ditawarkan', '')}
    """

    folium.Marker(
        location=[row["Latitude"], row["Longitude"]],
        popup=popup,
        tooltip=row["Nama Nasabah / PIC Usaha"],
        icon=folium.Icon(color="blue", icon="info-sign")
    ).add_to(marker_group)

# =====================================================
# Tambah Rute Jika Titik Awal & Tujuan Dipilih
# =====================================================
if start_point != "-" and end_point != "-":
    if start_point == end_point:
        st.sidebar.warning("Titik awal dan tujuan tidak boleh sama!")
    else:
        df_start = filtered_df[filtered_df["Nama Usaha"] == start_point].head(1)
        df_end = filtered_df[filtered_df["Nama Usaha"] == end_point].head(1)

        if df_start.empty or df_end.empty:
            st.sidebar.error("Titik awal atau tujuan tidak ditemukan pada data ter-filter.")
        else:
            lat1, lon1 = df_start["Latitude"].values[0], df_start["Longitude"].values[0]
            lat2, lon2 = df_end["Latitude"].values[0], df_end["Longitude"].values[0]

            route_coords = get_route_osrm(lat1, lon1, lat2, lon2)

            if route_coords:
                folium.PolyLine(
                    locations=route_coords,
                    color="red",
                    weight=5,
                    tooltip=f"Rute dari {start_point} ke {end_point}"
                ).add_to(m)

                # optionally add markers for start/end with different icon
                folium.Marker(
                    location=[lat1, lon1],
                    popup=f"Start: {start_point}",
                    icon=folium.Icon(color="green", icon="play")
                ).add_to(m)

                folium.Marker(
                    location=[lat2, lon2],
                    popup=f"End: {end_point}",
                    icon=folium.Icon(color="darkblue", icon="stop")
                ).add_to(m)
            else:
                st.sidebar.error("Gagal mengambil rute dari OSRM. Coba lagi atau cek koneksi.")

# =========================
# Search Bar (Leaflet Search)
# =========================
Search(
    layer=marker_group,
    search_label="tooltip",
    placeholder="🔍 Cari Nama Usaha...",
    collapsed=False
).add_to(m)

# Render Map
st_folium(m, width=1400, height=550)

# =========================
# Summary Metrics
# =========================
st.subheader("📊 Status Closing")

total_berhasil = (filtered_df["Status Closing"] == "Berhasil").sum()
total_callback = (filtered_df["Status Closing"] == "Callback").sum()
total_potensial = (filtered_df["Status Closing"] == "Potensial").sum()

col1, col2, col3 = st.columns(3)

col1.metric("Berhasil", total_berhasil)
col2.metric("Callback", total_callback)
col3.metric("Potensial", total_potensial)

# =========================
# Top 3 Produk Ditawarkan
# =========================
import plotly.express as px

# Hitung frekuensi produk
if "Produk yang Ditawarkan" in df.columns:
    product_counts = (
        df["Produk yang Ditawarkan"]
        .dropna()
        .str.split(", ")
        .explode()
        .value_counts()
    )
else:
    product_counts = pd.Series(dtype=int)

top3 = product_counts.head(3).reset_index()
top3.columns = ["Produk", "Jumlah"]

st.subheader("🏆 Top 3 Produk yang Paling Banyak Ditawarkan")

fig = px.bar(
    top3,
    x="Jumlah",
    y="Produk",
    orientation="h",
    text="Jumlah",
    title="Top 3 Produk",
)

# Styling seperti BI Tools
fig.update_layout(
    plot_bgcolor="rgba(0,0,0,0)",
    paper_bgcolor="rgba(0,0,0,0)",
    font=dict(size=16),
    xaxis_title="Jumlah Ditawarkan",
    yaxis_title="",
)

fig.update_traces(
    marker=dict(line=dict(width=1)),
    textposition="outside"
)

st.plotly_chart(fig, use_container_width=True)

# =========================
# Data Table
# =========================
st.subheader("📄 Data Visit")
# Kolom yang tidak ingin ditampilkan
drop_cols = ["Timestamp",
             "Koordinat (Latitude, Longitude), Contoh (-6.287825727208808, 107.0433026643262)",
             "Latitude",
             "Longitude"]

# Drop hanya kolom yang memang ada
display_df = filtered_df.drop(columns=[c for c in drop_cols if c in filtered_df.columns])

# Reset index agar mulai dari 1
display_df = display_df.reset_index(drop=True)   # reset index
display_df.index = display_df.index + 1          # mulai dari 1
display_df.index.name = "No"                     # kasih nama kolom index (opsional)

st.dataframe(display_df)
