import requests
from bs4 import BeautifulSoup
import pandas as pd
import re
from nltk.tokenize import word_tokenize, sent_tokenize

# Load custom stopwords
stopword_files = ["StopWords_Auditor.txt", "StopWords_currencies.txt", "StopWords_DatesandNumbers.txt" ,"StopWords_Generic.txt","StopWords_GenericLong.txt","StopWords_Geographic.txt",'StopWords_Names.txt']

# Combine all stopwords from the files
stop_words = set()
for file_name in stopword_files:
    with open(file_name, "r") as f:
        stop_words.update(f.read().splitlines())

# Load positive and negative word dictionaries
positive_words = set(open("positive-words.txt").read().split())
negative_words = set(open("negative-words.txt").read().split())

# Helper Functions
def clean_text(text):
    text = re.sub(r'[^\w\s]', '', text.lower())
    return [word for word in word_tokenize(text) if word not in stop_words]

def calculate_syllables(word):
    vowels = "aeiou"
    count = sum(1 for char in word if char in vowels)
    if word.endswith(("es", "ed")):
        count -= 1
    return max(count, 1)

def extract_text_from_url(url):
    try:
        response = requests.get(url)
        soup = BeautifulSoup(response.content, 'html.parser')
        title = soup.find('h1').get_text(strip=True)
        article_body = ' '.join(p.get_text(strip=True) for p in soup.find_all('p'))
        return title + "\n" + article_body
    except Exception as e:
        print(f"Failed to extract from {url}: {e}")
        return ""

# Main Processing
def analyze_articles(input_file, output_file):
    data = pd.read_excel(input_file)
    results = []

    for _, row in data.iterrows():
        url = row['URL']
        url_id = row['URL_ID']
        
        # Extract text
        text = extract_text_from_url(url)
        if not text:
            continue

        # Analysis
        tokens = clean_text(text)
        sentences = sent_tokenize(text)
        word_count = len(tokens)
        sentence_count = len(sentences)
        syllable_count = sum(calculate_syllables(word) for word in tokens)
        complex_words = sum(1 for word in tokens if calculate_syllables(word) > 2)

        positive_score = sum(1 for word in tokens if word in positive_words)
        negative_score = sum(1 for word in tokens if word in negative_words)
        polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
        subjectivity_score = (positive_score + negative_score) / (word_count + 0.000001)
        avg_sentence_length = word_count / sentence_count
        percentage_complex_words = complex_words / word_count
        fog_index = 0.4 * (avg_sentence_length + percentage_complex_words)
        avg_word_length = sum(len(word) for word in tokens) / word_count
        personal_pronouns = len(re.findall(r'\b(I|we|my|ours|us)\b', text, re.IGNORECASE))

        # Append result
        results.append({
            "URL_ID": url_id,
            "Positive Score": positive_score,
            "Negative Score": negative_score,
            "Polarity Score": polarity_score,
            "Subjectivity Score": subjectivity_score,
            "Avg Sentence Length": avg_sentence_length,
            "Percentage of Complex Words": percentage_complex_words,
            "Fog Index": fog_index,
            "Word Count": word_count,
            "Complex Word Count": complex_words,
            "Syllables per Word": syllable_count / word_count,
            "Personal Pronouns": personal_pronouns,
            "Avg Word Length": avg_word_length
        })

    # Save results to Excel
    output_df = pd.DataFrame(results)
    output_df.to_excel(output_file, index=False)

# Example Execution
analyze_articles("Input.xlsx", "Output Data Structure.xlsx")
