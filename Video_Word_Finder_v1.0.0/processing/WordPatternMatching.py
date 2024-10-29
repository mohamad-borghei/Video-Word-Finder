import  os
import pandas as pd
from itertools import combinations


def analyze_word_patterns(input_word, excel_file_path, output_folder):

    letters = list(input_word)
    word_length = len(letters)

    results = {i: [] for i in range(word_length, 2, -1)}

    df = pd.read_excel(excel_file_path)

    word_time_list = list(zip(df['word'], df.iloc[:, 1]))

    for pattern_length in range(word_length, 2, -1):
        num_patterns = word_length - pattern_length + 1

        for start_pos in range(num_patterns):
            pattern = ''.join(letters[start_pos:start_pos + pattern_length])

            for word, time in word_time_list:
                word_str = str(word)

                if word_str != input_word:
                    if len(word_str) >= pattern_length:
                        if word_str[:pattern_length] == pattern:
                            results[pattern_length].append((word, time))

    output_data = []

    for pattern_length in sorted(results.keys(), reverse=True):
        for word, time in results[pattern_length]:
            matching_pattern = str(word)[:pattern_length]
            output_data.append({
                'word': word,
                'time': time,
                'match_length': pattern_length,
                'matching_pattern': matching_pattern,
                'original_word': input_word
            })

    output_df = pd.DataFrame(output_data)

    if not output_df.empty:
        output_df = output_df[['word', 'time', 'match_length', 'matching_pattern', 'original_word']]

        output_file_path = os.path.join(output_folder, 'word_pattern_results.xlsx')
        output_df.to_excel(output_file_path, index=False)

    return results


def process_multiple_words(words_list, excel_file_path ,output_folder):

    all_results = {}
    for word in words_list:
        print(f"Processing the word '{word}' to find the most similar words to it.")
        results = analyze_word_patterns(word, excel_file_path, output_folder)
        all_results[word] = results
    return all_results


