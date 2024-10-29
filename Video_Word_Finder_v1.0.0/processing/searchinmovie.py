import speech_recognition as sr
import pandas as pd
from pydub import AudioSegment
import math
import os
from datetime import timedelta
from typing import Tuple, List, Dict
import logging

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def convert_seconds_to_timestamp(seconds: float) -> str:
    time_delta = timedelta(seconds=seconds)
    hours = time_delta.seconds // 3600
    minutes = (time_delta.seconds % 3600) // 60
    seconds = time_delta.seconds % 60

    return f"{hours:02d}:{minutes:02d}:{seconds:02d}" if hours > 0 else f"{minutes:02d}:{seconds:02d}"


def process_audio_chunk(
        recognizer: sr.Recognizer,
        chunk,
        offset: float = 0,
        chunk_duration: float = 30
) -> Tuple[List[Dict], List[Dict]]:
    try:
        result = recognizer.recognize_google(
            chunk,
            language="fa-IR,en-US",
            show_all=True
        )

        if not result or 'alternative' not in result:
            return [], []

        full_text = result['alternative'][0]['transcript']
        words = full_text.split()
        sentences = [{'sentence': s.strip()} for s in full_text.replace('؟', '.').replace('!', '.').split('.') if s.strip()]

        word_timestamps = []
        if words:
            time_per_word = chunk_duration / len(words)
            word_timestamps = [
                {
                    'word': word,
                    'time': convert_seconds_to_timestamp(offset + (i * time_per_word)),
                    'time_in_seconds': round(offset + (i * time_per_word), 2)
                }
                for i, word in enumerate(words)
            ]

        return word_timestamps, sentences

    except sr.UnknownValueError:
        logger.warning(f"Speech recognition error at second {offset}")
        return [], []
    except sr.RequestError as e:
        logger.error(f"Google service error at second {offset}: {e}")
        return [], []
    except Exception as e:
        logger.error(f"Unexpected error in audio processing: {str(e)}")
        return [], []


def process_audio_file(
        audio_path: str,
        chunk_duration: int = 30
) -> Tuple[List[Dict], List[Dict]]:

    temp_wav = None

    try:
        logger.info("Converting audio file...")
        audio = AudioSegment.from_file(audio_path)
        temp_wav = "temp_audio.wav"
        audio.export(temp_wav, format="wav")

        total_duration = len(audio) / 1000
        all_words = []
        all_sentences = []

        recognizer = sr.Recognizer()

        with sr.AudioFile(temp_wav) as source:
            recognizer.adjust_for_ambient_noise(source)

            logger.info(f"Starting to process {math.ceil(total_duration)} seconds of audio file...")

            for offset in range(0, math.ceil(total_duration), chunk_duration):
                remaining_duration = min(chunk_duration, total_duration - offset)
                logger.info(f"Processing section from {offset} to {offset + remaining_duration} seconds...")

                audio_chunk = recognizer.record(source, duration=remaining_duration)
                chunk_words, chunk_sentences = process_audio_chunk(
                    recognizer,
                    audio_chunk,
                    offset,
                    remaining_duration
                )
                all_words.extend(chunk_words)
                all_sentences.extend(chunk_sentences)

        return all_words, all_sentences

    finally:
        if temp_wav and os.path.exists(temp_wav):
            os.remove(temp_wav)


def save_excel_file(
        df: pd.DataFrame,
        output_path: str,
        sheet_name: str,
        column_widths: Dict[str, int]
) -> None:

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        worksheet = writer.sheets[sheet_name]
        for column, width in column_widths.items():
            worksheet.column_dimensions[column].width = width


def save_results(
        words: List[Dict],
        sentences: List[Dict],
        output_excel_words: str,
        output_excel_sentences: str,
        keyword: str = None
) -> None:

    try:
        if not words and not sentences:
            logger.warning("No text recognized!")
            return

        if words:
            df_words = pd.DataFrame(words)
            save_excel_file(
                df_words,
                output_excel_words,
                'Words',
                {'A': 15, 'B': 10, 'C': 10}
            )
            logger.info(f"Words Excel file saved at {output_excel_words}. Word count: {len(words)}")

        if sentences:
            df_sentences = pd.DataFrame(sentences)
            save_excel_file(
                df_sentences,
                output_excel_sentences,
                'Sentences',
                {'A': 100}
            )
            logger.info(f"Sentences Excel file saved at {output_excel_sentences}. Sentence count: {len(sentences)}")

        if keyword and words:
            save_keyword_results(df_words, keyword, output_excel_words.replace('.xlsx', '_keyword.xlsx'))

    except Exception as e:
        logger.error(f"Error in saving Excel files: {str(e)}")


def save_keyword_results(df_words: pd.DataFrame, keyword: str, output_path: str) -> None:

    keyword_df = df_words[df_words['word'].str.lower() == keyword.lower()]
    if not keyword_df.empty:
        save_excel_file(
            keyword_df,
            output_path,
            'Keyword',
            {'A': 30, 'B': 15, 'C': 15}
        )
        logger.info(f"Keyword '{keyword}' results saved at {output_path}. Occurrence count: {len(keyword_df)}")
    else:
        logger.info(f"Keyword '{keyword}' not found in the text.")


