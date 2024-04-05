import whisper
from pathlib import Path
import openpyxl
import time
import psutil  # Import psutil for CPU usage

def preprocess_transcription(transcription):
    # Replace symbols with spaces
    escape_symbols = ".,"
    for symbol in escape_symbols:
        transcription = transcription.replace(symbol, " ")
    
    words = transcription.split()
    word_to_num = {
        'ZERO': '0',
        'ONE': '1',
        'TWO': '2',
        'THREE': '3',
        'FOUR': '4',
        'FIVE': '5',
        'SIX': '6',
        'SEVEN': '7',
        'EIGHT': '8',
        'NINE': '9'
    }

    processed_words = []
    for word in words:
        if word.upper() in word_to_num:
            processed_words.append(word_to_num[word.upper()])
        else:
            processed_words.append(word)

    first_letters = [word[0] for word in processed_words if word]
    final_output = ''.join(first_letters).lower()

    return final_output

# Function to calculate Word Error Rate (WER)
def wer(ref, hyp):
    # Create a matrix to store distances
    d = [[0] * (len(hyp) + 1) for _ in range(len(ref) + 1)] #len of matrix
    # Initialize first row and column of the matrix
    for i in range(len(ref) + 1): 
        d[i][0] = i
    for j in range(len(hyp) + 1):
        d[0][j] = j
    # Populate the matrix
    for i in range(1, len(ref) + 1):
        for j in range(1, len(hyp) + 1):
            cost = 0 if ref[i - 1] == hyp[j - 1] else 1
            d[i][j] = min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + cost)
    # Calculate WER
    wer_value = float(d[len(ref)][len(hyp)]) / len(ref)
    return wer_value * 100  # Convert WER to percentage


# Load the model
model = whisper.load_model('large')

# Define the directory containing the audio files
directory = Path('C:/AudioCAPTCHASamples/CAPTCHAS.Net_AUDIO')

# Load reference transcriptions
reference_transcriptions = {}
for ref_file in Path('C:/AudioCAPTCHASamples/RT_CAPTCHAS.Net').glob('*.txt'):
    with open(ref_file, 'r') as f:
        reference_transcription = f.read().strip()
        reference_transcriptions[ref_file.stem] = reference_transcription
        print(f"Ground Truth for {ref_file.stem}: {reference_transcription}")

# Create a new Excel workbook
workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.title = "large Model"
sheet['A1'] = "Audio Sample"
sheet['B1'] = "Transcription Results"
sheet['C1'] = "Ground Truth"
sheet['D1'] = "Transcription Time (seconds)"
sheet['E1'] = "Confidence Level"
sheet['F1'] = "Language Identified"
sheet['G1'] = "WER (%)"
sheet['H1'] = "CPU Usage (%)"

# Get a sorted list of audio files in the directory, sorting numerically
audio_files = sorted(directory.glob('*.mp3'), key=lambda x: int(x.stem[3:]))

# Row index to keep track of where to write the results in the Excel sheet
row_index = 2

# Iterate over all audio files in the directory
for audio_file in audio_files:
    start_time = time.time()  # Record start time
    
    result = model.transcribe(str(audio_file), language='en', verbose=True)
    print(f"Transcription result: {result}")
    
    end_time = time.time()  # Record end time
    transcription_time = end_time - start_time  # Calculate transcription time
    print(f"Transcription time: {transcription_time} seconds")

    # Preprocess the transcription
    processed_transcription = preprocess_transcription(result['text'])
    print(f"Processed transcription: {processed_transcription}")

    # Calculate Word Error Rate (WER) using reference transcription
    reference_text = reference_transcriptions.get(audio_file.stem, '')
    print(f"Reference transcription: {reference_text}")

    if reference_text:
        # Calculate WER
        wer_percent = wer(reference_text, processed_transcription)
        print(f"WER (%): {wer_percent}")
        # Estimate confidence based on WER
        confidence_level = "High" if wer_percent < 20 else "Low"  # Adjust threshold as needed
    else:
        confidence_level = "Reference not available"
        wer_percent = None
    
    # Get CPU usage
    cpu_usage = psutil.cpu_percent()  # Extract CPU usage using psutil
    print(f"CPU Usage (%): {cpu_usage}")

    # Write the audio sample name, transcription results, ground truth, transcription time, confidence level, language identification, WER percentage, and CPU usage to the Excel sheet
    sheet[f'A{row_index}'] = audio_file.name
    sheet[f'B{row_index}'] = processed_transcription
    sheet[f'C{row_index}'] = reference_text
    sheet[f'D{row_index}'] = transcription_time
    sheet[f'E{row_index}'] = confidence_level
    sheet[f'F{row_index}'] = result['language'] if 'language' in result else "Language not identified"
    sheet[f'G{row_index}'] = wer_percent if wer_percent is not None else "N/A"
    sheet[f'H{row_index}'] = cpu_usage

    # Increment row index for next entry
    row_index += 1

# Save the Excel workbook
workbook.save("CAPTranscriptions.xlsx")

