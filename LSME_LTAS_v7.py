import os
import subprocess
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy.signal import savgol_filter
from scipy.stats import ttest_ind

def run_praat_script(praat_path, script_content, wav_file_path, output_ltas_path):
    """ Führt ein Praat-Skript aus, um LTAS-Daten zu generieren """
    praat_script_path = "temp_script.praat"
    script_content = script_content.replace("INPUT_WAV_PATH", wav_file_path)
    script_content = script_content.replace("OUTPUT_LTAS_PATH", output_ltas_path)
    try:
        with open(praat_script_path, "w", encoding="utf-8") as script_file:
            script_file.write(script_content)
        subprocess.run([praat_path, "--run", praat_script_path], check=True)
    except subprocess.CalledProcessError as e:
        print(f"Fehler bei der Ausführung von Praat: {e}")
    finally:
        if os.path.exists(praat_script_path):
            os.remove(praat_script_path)

def read_ltas_file(file_path):
    """ Liest eine LTAS-Datei ein und gibt Frequenzen und Amplituden zurück """
    frequencies, amplitudes = [], []
    try:
        with open(file_path, "r", encoding="utf-8") as file:
            data_lines = file.readlines()[13:]  # Kopfzeilen überspringen
            base_frequency, bin_step = 5, 10
            for i, amplitude_value in enumerate(data_lines):
                try:
                    amplitude = float(amplitude_value.strip())
                    frequencies.append(base_frequency + i * bin_step)
                    amplitudes.append(amplitude)
                except ValueError:
                    continue
    except Exception as e:
        print(f"Fehler beim Verarbeiten der Datei {file_path}: {e}")
    return np.array(frequencies), np.array(amplitudes)

def normalize_amplitudes(all_amplitudes):
    """ Normalisiert alle Amplituden auf einen globalen Mittelwert """
    global_mean = np.mean([np.mean(amp) for amp in all_amplitudes])  # Globaler Mittelwert über ALLE Messungen
    return [amp - np.mean(amp) + global_mean for amp in all_amplitudes]  # Alle Messungen auf diesen Mittelwert angleichen

def get_user_selected_folders(base_directory):
    """Fragt den Benutzer ab, welche Unterordner verarbeitet werden sollen"""
    subfolders = [f.name for f in os.scandir(base_directory) if f.is_dir()]
    selected = []

    print("\nVerfügbare Unterordner:")
    for folder in subfolders:
        response = input(f'  Soll "{folder}" verarbeitet werden? (y/n): ').strip().lower()
        if response == 'y':
            selected.append(folder)

    print(f"\nAusgewählte Ordner: {selected if selected else 'Keine'}")
    return selected


def plot_all_ltas(ltas_data, p_values=[]):
    """Erstellt einen Plot für die mittleren LTAS-Werte aller Unterordner mit Standardabweichung und geglätteten p-Werten"""
    fig, ax1 = plt.subplots(figsize=(12, 5.5))

    # Glättung der p-Werte mit demselben Filter wie für die Amplituden
    if len(p_values) > 0:
        smoothed_p = savgol_filter(p_values, 37, 2)
        smoothed_p = np.clip(smoothed_p, 1e-10, 1)  # Vermeidet log(0)

    # Plot für LTAS-Daten
    for folder_name, (frequencies, avg_amplitudes, std_amplitudes, file_count) in ltas_data.items():
        smoothed_avg = savgol_filter(avg_amplitudes, 37, 2)
        smoothed_std = savgol_filter(std_amplitudes, 37, 2)

        ax1.plot(frequencies, smoothed_avg, linewidth=2, label=f'{folder_name} (n={file_count})')
        ax1.fill_between(frequencies, smoothed_avg - smoothed_std, smoothed_avg + smoothed_std, alpha=0.3)

    ax1.set_xscale('log')
    ax1.set_title("LTAS Nordwind und Sonne", fontsize=14)
    ax1.set_xlabel("Frequenz (Hz)", fontsize=12)
    ax1.set_ylabel("Amplitude (dB)", fontsize=12)
    #ax1.set_xlim(1500, 5000)
    ax1.set_xlim(200, 5000)
    #ticks = [1500, 2000, 3000, 4000, 5000]
    ticks = [200, 400, 600 , 800, 1000, 2000, 3000, 4000, 5000]
    ax1.set_xticks(ticks)
    ax1.set_xticklabels([str(t) for t in ticks])
    ax1.set_ylim(5, 40)
    ax1.grid(True, which='both', linestyle='--', alpha=0.5)
    ax1.legend(loc='upper left')
    ax1.axvspan(2250, 3880, color='violet', alpha=0.25)

    # Plot für p-Werte
    if len(p_values) > 0:
        ax2 = ax1.twinx()
        ax2.plot(frequencies, smoothed_p, linestyle='dashed', color='black', label='p-Wert (geglättet)')
        ax2.axhline(y=0.05, color='red', linestyle='dotted', label='Signifikanzniveau 0.05')
        ax2.set_yscale('log')
        ax2.set_ylabel("p-Wert")
        ax2.set_ylim(1e-4, 1)
        ax2.legend(loc='upper right')

    plt.tight_layout()
    plt.show()

def cleanup_temp_files(base_directory):
    """ Löscht alle temporären LTAS-Dateien nach der Verarbeitung """
    for root, _, files in os.walk(base_directory):
        for file in files:
            if file.endswith(".Ltas"):
                file_path = os.path.join(root, file)
                try:
                    os.remove(file_path)
                except Exception as e:
                    print(f"Fehler beim Löschen von {file_path}: {e}")

if __name__ == "__main__":
    praat_executable_path = "C:/Program Files/praat6426_win-intel64/Praat.exe"
    base_directory = r"D:\LSME\LSME Auswertung\Data\LTAS Stilvergleich"
    output_excel_path = os.path.join (base_directory, "LTAS Output.xlsx")

    # Benutzerauswahl der Ordner
    selected_subdirs = get_user_selected_folders(base_directory)
    if not selected_subdirs:
        print("Abbruch: Keine Ordner ausgewählt.")
        exit()

    praat_script = """
    Read from file: "INPUT_WAV_PATH"
    ltasObject = To Ltas: 10
    Save as short text file: "OUTPUT_LTAS_PATH"
    Remove
    """

    all_ltas_data = {}
    all_amplitudes = []
    all_filenames = []
    all_frequencies = None

    # 1. **Verarbeitung der ausgewählten Ordner**
    for subdir in selected_subdirs:
        subdir_path = os.path.join(base_directory, subdir)

        for file_name in os.listdir(subdir_path):
            if file_name.lower().endswith(".wav"):
                input_wav_path = os.path.join(subdir_path, file_name)
                output_ltas_path = os.path.join(subdir_path, f"{os.path.splitext(file_name)[0]}.Ltas")

                run_praat_script(praat_executable_path, praat_script, input_wav_path, output_ltas_path)
                frequencies, amplitudes = read_ltas_file(output_ltas_path)

                if frequencies.size > 0 and amplitudes.size > 0:
                    if all_frequencies is None:
                        all_frequencies = frequencies  # Frequenzen für alle gleich
                    all_amplitudes.append(amplitudes)
                    all_filenames.append((subdir, file_name))

    # 2. **Globale Normalisierung auf einen gemeinsamen Mittelwert**
    if all_amplitudes:
        normalized_amplitudes = normalize_amplitudes(all_amplitudes)

        # 3. **Excel-Datei erstellen mit allen normalisierten Daten**
        with pd.ExcelWriter(output_excel_path) as writer:
            folder_data = {}

            for (subdir, file_name), normalized_amp in zip(all_filenames, normalized_amplitudes):
                if subdir not in folder_data:
                    folder_data[subdir] = {"Frequenz (Hz)": all_frequencies}

                folder_data[subdir][file_name] = normalized_amp

            for subdir, data in folder_data.items():
                df = pd.DataFrame(data)
                df.to_excel(writer, sheet_name=subdir, index=False)
                print(f"Ergebnisse für {subdir} gespeichert.")

            print(f"Gesamtauswertung gespeichert unter: {output_excel_path}")

        # Berechnung von Mittelwert, Standardabweichung und p-Werten
        folder_keys = list(folder_data.keys())
        if len(folder_keys) == 2:
            amplitudes_1 = np.array(
                [folder_data[folder_keys[0]][key] for key in folder_data[folder_keys[0]] if key != "Frequenz (Hz)"])
            amplitudes_2 = np.array(
                [folder_data[folder_keys[1]][key] for key in folder_data[folder_keys[1]] if key != "Frequenz (Hz)"])

            mean_amplitude_1 = np.mean(amplitudes_1, axis=0)
            std_amplitude_1 = np.std(amplitudes_1, axis=0)
            mean_amplitude_2 = np.mean(amplitudes_2, axis=0)
            std_amplitude_2 = np.std(amplitudes_2, axis=0)
            file_count_1 = len(amplitudes_1)
            file_count_2 = len(amplitudes_2)

            p_values = np.array([ttest_ind(amplitudes_1[:, i], amplitudes_2[:, i], equal_var=False)[1] for i in
                                 range(amplitudes_1.shape[1])])

            all_ltas_data[folder_keys[0]] = (all_frequencies, mean_amplitude_1, std_amplitude_1, file_count_1)
            all_ltas_data[folder_keys[1]] = (all_frequencies, mean_amplitude_2, std_amplitude_2, file_count_2)

            plot_all_ltas(all_ltas_data, p_values)
        else:
            for subdir, data in folder_data.items():
                amplitudes = np.array([data[key] for key in data if key != "Frequenz (Hz)"])
                mean_amplitude = np.mean(amplitudes, axis=0)
                std_amplitude = np.std(amplitudes, axis=0)
                file_count = len(amplitudes)

                all_ltas_data[subdir] = (all_frequencies, mean_amplitude, std_amplitude, file_count)

        if all_ltas_data:
            plot_all_ltas(all_ltas_data)

        # 5. **Löschen der temporären LTAS-Dateien**
        cleanup_temp_files(base_directory)