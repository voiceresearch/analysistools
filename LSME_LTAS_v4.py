import os
import subprocess
import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
from scipy.interpolate import CubicSpline
from scipy.signal import savgol_filter


def run_praat_script(praat_path, script_content, wav_file_path, output_ltas_path):
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
    frequencies = []
    amplitudes = []
    try:
        with open(file_path, "r", encoding="utf-8") as file:
            data_lines = file.readlines()[13:]
            base_frequency = 5  # Korrektur: Erste Frequenz bei 5 Hz
            bin_step = 10
            for i, amplitude_value in enumerate(data_lines):
                try:
                    amplitude = float(amplitude_value.strip())
                    frequencies.append(base_frequency + i * bin_step)
                    amplitudes.append(amplitude)
                except ValueError:
                    continue
    except Exception as e:
        print(f"Fehler beim Verarbeiten der Datei: {e}")
    return frequencies, amplitudes


def normalize_and_filter(frequencies, amplitudes, target_mean):
    freq_array = np.array(frequencies)
    amp_array = np.array(amplitudes)

    # Filter auf 20-10000 Hz
    mask = (freq_array >= 20) & (freq_array <= 10000)
    filtered_freq = freq_array[mask]
    filtered_amp = amp_array[mask]

    # Normalisierung
    current_mean = np.mean(filtered_amp)
    if current_mean != 0:
        filtered_amp *= target_mean / current_mean

    return filtered_freq.tolist(), filtered_amp.tolist()


def process_and_smooth(frequencies, amplitudes):
    # Interpolation
    spline = CubicSpline(frequencies, amplitudes)
    fine_freq = np.logspace(np.log10(20), np.log10(10000), 500)
    fine_amp = spline(fine_freq)

    # Glättung
    return fine_freq, savgol_filter(fine_amp, window_length=21, polyorder=3)


def plot_average_ltas(averaged_data):
    plt.figure(figsize=(12, 6))

    # Verarbeitung der gemittelten Daten
    fine_freq, smooth_amp = process_and_smooth(*averaged_data)

    plt.plot(fine_freq, smooth_amp,
             linewidth=2,
             color='#2a9d8f',
             label='Gemitteltes LTAS')

    plt.xscale('log')
    plt.title("Gemitteltes Long-Term Average Spectrum (20-10,000 Hz)", fontsize=14)
    plt.xlabel("Frequenz (Hz)", fontsize=12)
    plt.ylabel("Amplitude (dB)", fontsize=12)

    plt.xlim(50, 10000)
    plt.ylim(-20, 20)
    ticks = [20, 50, 100, 200, 500, 1000, 2000, 5000, 10000]
    plt.xticks(ticks, [f'{tick}' for tick in ticks], rotation=45)
    plt.grid(True, which='both', linestyle='--', alpha=0.5)
    plt.legend(loc='upper right')
    plt.tight_layout()
    plt.show()


def calculate_average_ltas(normalized_ltas):
    # Sammle alle Amplitudenwerte
    all_amps = []
    ref_freq = None

    for freq, amp in normalized_ltas:
        if ref_freq is None:
            ref_freq = freq
        all_amps.append(amp)

    # Berechne Mittelwerte
    mean_amps = np.mean(all_amps, axis=0)
    return ref_freq, mean_amps


if __name__ == "__main__":
    praat_executable_path = "C:/Program Files/praat6426_win-intel64/Praat.exe"
    input_directory = "C:/Users/jerem/Desktop/Studienassistenz/LSME/LTAS Nus Test"
    output_directory = "C:/Users/jerem/Desktop/Studienassistenz/LSME/LTAS Nus Test"
    os.makedirs(output_directory, exist_ok=True)

    praat_script = """
    Read from file: "INPUT_WAV_PATH"
    ltasObject = To Ltas: 10
    Save as short text file: "OUTPUT_LTAS_PATH"
    Remove
    """

    all_ltas = []
    for file_name in os.listdir(input_directory):
        if file_name.lower().endswith(".wav"):
            input_wav_path = os.path.join(input_directory, file_name)
            output_ltas_path = os.path.join(output_directory, f"{os.path.splitext(file_name)[0]}.Ltas")

            run_praat_script(praat_executable_path, praat_script, input_wav_path, output_ltas_path)
            frequencies, amplitudes = read_ltas_file(output_ltas_path)

            if frequencies and amplitudes:
                all_ltas.append((frequencies, amplitudes))

    if all_ltas:
        # Normalisierung und Filterung
        normalized_data = []
        mean_values = []

        # Erste Berechnung der Mittelwerte
        for freq, amp in all_ltas:
            filtered_freq, filtered_amp = normalize_and_filter(freq, amp, 1)
            mean_values.append(np.mean(filtered_amp))

        target_mean = np.mean(mean_values)

        # Finale Normalisierung
        for freq, amp in all_ltas:
            filtered_freq, filtered_amp = normalize_and_filter(freq, amp, target_mean)
            normalized_data.append((filtered_freq, filtered_amp))

        # Mittelwertbildung
        averaged_freq, averaged_amp = calculate_average_ltas(normalized_data)

        # Visualisierung und Speicherung
        plot_average_ltas((averaged_freq, averaged_amp))

        # Excel-Export
        df = pd.DataFrame({
            'Frequency (Hz)': averaged_freq,
            'Average Amplitude (dB)': averaged_amp
        })
        excel_path = os.path.join(output_directory, "Gemitteltes_LTAS.xlsx")
        df.to_excel(excel_path, index=False)
        print(f"Ergebnisse gespeichert in: {excel_path}")