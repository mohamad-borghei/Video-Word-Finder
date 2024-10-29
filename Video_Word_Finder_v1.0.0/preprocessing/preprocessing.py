import re

def clean_word(word):
    cleaned_word = re.sub(r'[^a-zA-Z\u0600-\u06FF]', '', word)
    return cleaned_word

def keep_only_numbers(number):
    cleaned_number = re.sub(r'[^0-9]', '', number)
    return cleaned_number






