import os
import string
from collections import Counter
import textract
from PyPDF2 import PdfReader
from docx import Document

def count_words(file_path):
    _, file_extension = os.path.splitext(file_path)

    if file_extension == ".pdf":
        return count_words_pdf(file_path)
    elif file_extension == ".docx":
        return count_words_docx(file_path)
    else:
        return count_words_text(file_path)

def count_words_pdf(file_path):
    with open(file_path, 'rb') as file:
        pdf = PdfReader(file)
        word_count = Counter()
        for page in pdf.pages:
            text = page.extract_text()
            word_count.update(text.split())
        return word_count

def count_words_docx(file_path):
    doc = Document(file_path)
    text = '\n'.join([paragraph.text for paragraph in doc.paragraphs])
    return Counter(text.split())

def count_words_text(file_path):
    with open(file_path, 'r') as file:
        text = file.read()
        text = text.translate(str.maketrans("", "", string.punctuation))
        words = text.split()
        return Counter(words)

def analyze_common_words(word_count, num_common_words):
    common_words = word_count.most_common(num_common_words)
    print("Most common words:")
    for word, count in common_words:
        print(word, ":", count)

def search_word_frequency(word_count, search_word):
    word_frequency = word_count.get(search_word, 0)
    print("Frequency of the word", search_word, ":", word_frequency)

def main():
    print("Welcome to the File Word Count Analyzer!")
    file_path = input("Enter the path to the file: ")

    if os.path.isfile(file_path):
        word_count = count_words(file_path)
        print("Total words in the file:", sum(word_count.values()))

        while True:
            choice = input("Do you want to:\n1. View most common words\n2. Search for a specific word\n3. Exit\nEnter your choice (1, 2, or 3): ")

            if choice == '1':
                num_common_words = int(input("Enter the number of most common words to display: "))
                analyze_common_words(word_count, num_common_words)
                print()
            elif choice == '2':
                search_word = input("Enter the word to search: ")
                search_word_frequency(word_count, search_word)
                print()
            elif choice == '3':
                break
            else:
                print("Invalid choice. Please try again.")

    else:
        print("Error: File not found.")

    print("Thank you for using the File Word Count Analyzer!")

if __name__ == "__main__":
    main()

