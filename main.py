import re
import pandas as pd
from nltk.tokenize import sent_tokenize, word_tokenize
from newspaper import Article, ArticleException


# load stop variables from txt files and combine them into one set
def load_stop_words(stopword_files):
    stop_words = set()
    for file in stopword_files:
        with open(file, 'r') as f:
            stop_words.update(word.strip() for word in f.readlines())
    return stop_words


# Define function to extract article text from URL
def extract_article_text(url):
    # Get title and article content from URL using Newspaper3k library
    try:
        article = Article(url)
        article.download()
        article.parse()
        article_title = article.title
        article_text = article.text
        return article_title, article_text

    except ArticleException as e:
        print(f"Error: Unable to fetch article from {url}. Reason: {e}")
        return None, None


def count_syllables(word):
    vowels = "aeiouy"
    word = word.lower()
    count = 0
    if word[0] in vowels:
        count += 1
    for index in range(1, len(word)):
        if word[index] in vowels and word[index - 1] not in vowels:
            count += 1
    if word.endswith("e"):
        count -= 1
    if count == 0:
        count += 1
    return count


def analyze_text(text, stop_words, positive_words, negative_words):
    # Clean and tokenize text
    tokens = word_tokenize(text)
    tokens = [token.lower() for token in tokens if token.isalnum() and token.lower() not in stop_words]
    sentences = sent_tokenize(text)

    # Compute all required variables from doc
    positive_score = sum(1 for token in tokens if token in positive_words)
    negative_score = sum(1 for token in tokens if token in negative_words)
    total_words = len(tokens)

    polarity_score = (positive_score - negative_score) / (positive_score + negative_score + 0.000001)
    subjectivity_score = (positive_score + negative_score) / (total_words + 0.000001)

    try:
        avg_sentence_length = total_words / len(sentences)
    except ZeroDivisionError:
        avg_sentence_length = 0

    try:
        avg_word_length = sum(len(token) for token in tokens) / total_words
    except ZeroDivisionError:
        avg_word_length = 0

    complex_word_count = sum(1 for token in tokens if count_syllables(token) > 2)
    percentage_complex_words = (complex_word_count / total_words) * 100
    fog_index = 0.4 * (avg_sentence_length + percentage_complex_words)

    personal_pronouns = sum(1 for token in tokens if re.match(r'\b(i|we|my|ours|us)\b', token))

    syllable_count = sum(count_syllables(token) for token in tokens)
    syllable_per_word = syllable_count / total_words

    return {
        'POSITIVE SCORE': positive_score,
        'NEGATIVE SCORE': negative_score,
        'POLARITY SCORE': polarity_score,
        'SUBJECTIVITY SCORE': subjectivity_score,
        'AVG SENTENCE LENGTH': avg_sentence_length,
        'PERCENTAGE OF COMPLEX WORDS': percentage_complex_words,
        'FOG INDEX': fog_index,
        'AVG NUMBER OF WORDS PER SENTENCE': avg_sentence_length,
        'COMPLEX WORD COUNT': complex_word_count,
        'WORD COUNT': total_words,
        'SYLLABLE PER WORD': syllable_per_word,
        'PERSONAL PRONOUNS': personal_pronouns,
        'AVG WORD LENGTH': avg_word_length
    }


def main():
    # Load input URLs from the Input.xlsx file
    input_df = pd.read_excel('Input.xlsx')
    urls = input_df['URL'].tolist()

    # Load stop words
    stopword_files = [
        './resource/StopWords/StopWords_Auditor.txt',
        './resource/StopWords/StopWords_Currencies.txt',
        './resource/StopWords/StopWords_DatesandNumbers.txt',
        './resource/StopWords/StopWords_Generic.txt',
        './resource/StopWords/StopWords_GenericLong.txt',
        './resource/StopWords/StopWords_Geographic.txt',
        './resource/StopWords/StopWords_Names.txt'
    ]

    stop_words = load_stop_words(stopword_files)

    # Load Positive and Negative words
    with open('./resource/MasterDictionary/positive_words.txt') as f:
        positive_words = set(word.strip() for word in f.readlines())
    with open('./resource/MasterDictionary/negative_words.txt') as f:
        negative_words = set(word.strip() for word in f.readlines())

    # Initialize the output DataFrame
    output_df = input_df.copy()

    # Analyze each URL
    for index, url in enumerate(urls):
        print(f'Processing URL {index + 1}/{len(urls)}: {url}')
        article_text = extract_article_text(url)[1]
        if article_text:
            with open(f'{index}_article_text.txt', 'w', encoding='utf-8') as f:
                f.write(article_text)
            analysis_results = analyze_text(article_text, stop_words, positive_words, negative_words)

            # Update the output DataFrame with the analysis results
            for key, value in analysis_results.items():
                output_df.at[index, key] = value

    # Save output DataFrame to Output.xlsx
    output_df.to_excel('Output.xlsx', index=False)
    print('Output saved to Output.xlsx')


if __name__ == '__main__':
    main()
