import pandas as pd
import os
import matplotlib.pyplot as plt
import numpy as np
from scipy.stats import ttest_ind

# ---------------------------
# 1. Daten einlesen und verarbeiten
# ---------------------------
folder_path = r"C:\Users\jerem\Filr\Meine Dateien\Projekte\LSME\Data\VRP XML"
input_file = os.path.join(folder_path, "LSME_VRP_EDIT1.xlsx")
output_file = r"C:\Users\jerem\Desktop\Studienassistenz\LSME\LSME_VRP_EDIT2.xlsx"

# Excel-Daten laden
vrp_sheet = pd.read_excel(input_file, sheet_name="VRP", header=None)
data = vrp_sheet.iloc[1:]
headers = vrp_sheet.iloc[0].tolist()
messreihe_labels = vrp_sheet.iloc[2].tolist()


# Gruppen erstellen
def create_group_df(group_letter):
    cols = [i for i, label in enumerate(messreihe_labels) if label == group_letter]
    df = data.iloc[:, cols]
    df.columns = [headers[i] for i in cols]
    df.insert(0, 'Proband_in', data.iloc[:, 0])
    df.name = group_letter  # Für Debug-Zwecke
    return df


df_A = create_group_df('A')
df_B = create_group_df('B')
df_C = create_group_df('C')


# ---------------------------
# 2. Statistische Auswertung
# ---------------------------
import pandas as pd
import os
import matplotlib.pyplot as plt
import numpy as np
from scipy.stats import ttest_ind


# ---------------------------
# 1. Neue Kategorien für die Analyse
# ---------------------------
def extract_extended_stats(df):
    stats = {
        'sprech_dynamik': None, 'sprech_dynamik_std': None,
        'sprech_tonhoehe': None, 'sprech_tonhoehe_std': None,
        'gesangs_dynamik': None, 'gesangs_dynamik_std': None,
        'gesangs_tonhoehe': None, 'gesangs_tonhoehe_std': None,
        'norm_profile_coverage': None, 'norm_profile_coverage_std': None
    }

    def get_values(row_name):
        row = df[df['Proband_in'] == row_name]
        if not row.empty:
            values = row.iloc[0, 1:].dropna()
            try:
                return values.astype(float)
            except:
                return pd.Series(dtype=float)
        return pd.Series(dtype=float)

    # Daten extrahieren
    stats['sprech_dynamik'] = get_values('Sprech-Dynamikumfang')
    stats['sprech_dynamik_std'] = stats['sprech_dynamik'].std()
    stats['sprech_tonhoehe'] = get_values('Sprech-Tonumfang')
    stats['sprech_tonhoehe_std'] = stats['sprech_tonhoehe'].std()
    stats['gesangs_dynamik'] = get_values('Gesangs-Dynamikumfang')
    stats['gesangs_dynamik_std'] = stats['gesangs_dynamik'].std()
    stats['gesangs_tonhoehe'] = get_values('Gesangs-Tonumfang')
    stats['gesangs_tonhoehe_std'] = stats['gesangs_tonhoehe'].std()
    stats['norm_profile_coverage'] = get_values('norm_profile_coverage')
    stats['norm_profile_coverage_std'] = stats['norm_profile_coverage'].std()

    return stats


# Berechnungen für Gruppen
stats_A = extract_extended_stats(df_A)
stats_B = extract_extended_stats(df_B)
stats_C = extract_extended_stats(df_C)


# ---------------------------
# 2. p-Werte berechnen
# ---------------------------
def calculate_p_values(values1, values2):
    if len(values1) > 1 and len(values2) > 1:
        _, p_value = ttest_ind(values1, values2, equal_var=False)
        return format(round(p_value, 20), ".20f")  # Korrekt gerundet auf drei Nachkommastellen
    return None

p_values_B_A = [calculate_p_values(stats_A[k], stats_B[k]) for k in
                ['sprech_dynamik', 'sprech_tonhoehe', 'gesangs_dynamik', 'gesangs_tonhoehe', 'norm_profile_coverage']]
p_values_C_B = [calculate_p_values(stats_B[k], stats_C[k]) for k in
                ['sprech_dynamik', 'sprech_tonhoehe', 'gesangs_dynamik', 'gesangs_tonhoehe', 'norm_profile_coverage']]


# ---------------------------
# 3. Einzelne Diagramme plotten mit A-B-C Reihenfolge
# ---------------------------
def plot_single_bar(category, values_A, values_B, values_C, errors_A, errors_B, errors_C, p_value_B_A, p_value_C_B,
                    n_A, n_B, n_C, title):
    x = np.array([f"A (n={n_A})", f"B (n={n_B})", f"C (n={n_C})"])
    values = [np.mean(values_A), np.mean(values_B), np.mean(values_C)]
    errors = [errors_A, errors_B, errors_C]

    fig, ax = plt.subplots(figsize=(8, 5))
    bars = ax.barh(x[::-1], values[::-1], xerr=errors[::-1], color=['skyblue', 'lightgreen', 'salmon'], edgecolor='black')

    # p-Werte einfügen
    if p_value_B_A is not None:
        ax.text(ax.get_xlim()[1] - 0.02,  # Rechtsbündig am Rand des Diagramms
                bars[1].get_y() + bars[1].get_height() / 2,  # Mittig auf dem Balken von B
                f"p={p_value_B_A}", va='center', ha='right', fontsize=10, color='black')

    if p_value_C_B is not None:
        ax.text(ax.get_xlim()[1] - 0.02,  # Rechtsbündig am Rand des Diagramms
                bars[0].get_y() + bars[0].get_height() / 2,  # Mittig auf dem Balken von C
                f"p={p_value_C_B}", va='center', ha='right', fontsize=10, color='black')

    if "Tonumfang" in title:
        ax.set_xlabel("Tonumfang [Halbtöne]")

    elif "Dynamikumfang" in title:
        ax.set_xlabel("Dynamikumfang [dB]")

    elif "Norm Profile Coverage" in title:
        ax.set_xlabel("Normstimmfeldabdeckung [%]")
    ax.set_title(title)
    ax.grid(axis='x', linestyle='--', alpha=0.7)
    plt.tight_layout()
    plt.show()


# Diagramme für jede Kategorie einzeln plotten
plot_single_bar("Sprech-Tonumfang", stats_A['sprech_tonhoehe'], stats_B['sprech_tonhoehe'], stats_C['sprech_tonhoehe'],
                stats_A['sprech_tonhoehe_std'], stats_B['sprech_tonhoehe_std'], stats_C['sprech_tonhoehe_std'],
                p_values_B_A[1], p_values_C_B[1], len(stats_A['sprech_tonhoehe']), len(stats_B['sprech_tonhoehe']), len(stats_C['sprech_tonhoehe']), "Sprech-Tonumfang")

plot_single_bar("Gesangs-Tonumfang", stats_A['gesangs_tonhoehe'], stats_B['gesangs_tonhoehe'],
                stats_C['gesangs_tonhoehe'], stats_A['gesangs_tonhoehe_std'], stats_B['gesangs_tonhoehe_std'],
                stats_C['gesangs_tonhoehe_std'], p_values_B_A[3], p_values_C_B[3], len(stats_A['gesangs_tonhoehe']), len(stats_B['gesangs_tonhoehe']), len(stats_C['gesangs_tonhoehe']), "Gesangs-Tonumfang")

plot_single_bar("Sprech-Dynamikumfang", stats_A['sprech_dynamik'], stats_B['sprech_dynamik'], stats_C['sprech_dynamik'],
                stats_A['sprech_dynamik_std'], stats_B['sprech_dynamik_std'], stats_C['sprech_dynamik_std'],
                p_values_B_A[0], p_values_C_B[0], len(stats_A['sprech_dynamik']), len(stats_B['sprech_dynamik']), len(stats_C['sprech_dynamik']), "Sprech-Dynamikumfang")

plot_single_bar("Gesangs-Dynamikumfang", stats_A['gesangs_dynamik'], stats_B['gesangs_dynamik'],
                stats_C['gesangs_dynamik'], stats_A['gesangs_dynamik_std'], stats_B['gesangs_dynamik_std'],
                stats_C['gesangs_dynamik_std'], p_values_B_A[2], p_values_C_B[2], len(stats_A['gesangs_dynamik']), len(stats_B['gesangs_dynamik']), len(stats_C['gesangs_dynamik']), "Gesangs-Dynamikumfang")

plot_single_bar("Norm Profile Coverage", stats_A['norm_profile_coverage'], stats_B['norm_profile_coverage'],
                stats_C['norm_profile_coverage'], stats_A['norm_profile_coverage_std'],
                stats_B['norm_profile_coverage_std'], stats_C['norm_profile_coverage_std'], p_values_B_A[4],
                p_values_C_B[4], len(stats_A['norm_profile_coverage']), len(stats_B['norm_profile_coverage']), len(stats_C['norm_profile_coverage']), "Norm Profile Coverage")
