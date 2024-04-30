"""Backend Coding Test"""
import os
import re as regex
from pathlib import Path
from typing import Any, List, Tuple

import xlsxwriter as excel
from nltk.probability import FreqDist
from nltk.tokenize import RegexpTokenizer
from nltk.tokenize import sent_tokenize as st


def create_excel_workbook(
    final_word_list: List[tuple],
    excel_file_name: str = "InterestingWords.xlsx",
) -> None:
    if ".xlsx" not in excel_file_name.lower():
        excel_file_name += ".xlsx"

    workbook = excel.Workbook(excel_file_name)
    worksheet = workbook.add_worksheet("Results")

    wrap_format = workbook.add_format(
        {"bold": True, "text_wrap": True, "valign": "top"}
    )

    cell_format = workbook.add_format({"text_wrap": True, "valign": "top"})

    worksheet.set_column("A1:C3", 25, wrap_format)

    final_row = len(final_word_list) + 1
    table_size = "A1:C{0}".format(final_row)
    worksheet.add_table(
        table_size,
        {
            "columns": [
                {"header": "Word (Total Occurrences)", "header_format": wrap_format},
                {"header": "Documents", "header_format": wrap_format},
                {
                    "header": "Sentences Containing the word",
                    "header_format": wrap_format,
                },
            ]
        },
    )

    for row, items in enumerate(final_word_list, 2):
        interesting_word = items[0][0]
        frequency = items[0][1]
        doc_name = items[1]
        sentences = items[2]

        word = f"{interesting_word} ({frequency})"
        word_col = "A{0}".format(row)
        doc_col = "B{0}".format(row)
        sent_col = "C{0}".format(row)

        worksheet.write(word_col, word, cell_format)
        worksheet.write(doc_col, doc_name, cell_format)
        worksheet.write(sent_col, sentences, cell_format)

    workbook.close()


def get_files_to_analyze(path_to_files: Path) -> List[Path]:
    """Get a list of text file paths from the given directory"""

    if not path_to_files.is_dir():
        raise FileNotFoundError(f"{path_to_files} is not a directory.")

    filenames: List[Path] | None = list(path_to_files.glob("*.txt"))

    if not filenames:
        raise FileNotFoundError(f"No .txt files in {path_to_files} directory.")

    return filenames


def generate_final_word_list(
    top_five_words: List[Tuple[Any, int]],
    filename: str,
    tokenized_sentences: List[str],
) -> List[tuple]:
    final_word_list = []
    new_line_separator = "\n\n"

    for word in top_five_words:
        sentences_list = get_sentences_that_contains_given_word(
            tokenized_sentences, word[0]
        )
        sentences = new_line_separator.join(sentences_list)
        final_word_list.append((word, filename, sentences))

    return final_word_list


def get_tokenized_list(document_text: str, stopwords: List[str]) -> List[str]:
    """
    Get a list of only alphanumeric and possessive words in tokenized form
    """
    regex_tk = RegexpTokenizer(r"\w+", gaps=False)
    tokenized_list = regex_tk.tokenize(document_text)

    tk_list_without_stopwords = [
        token.lower() for token in tokenized_list if token.lower() not in stopwords
    ]

    return tk_list_without_stopwords


# returns 3 sentences per word
def get_sentences_that_contains_given_word(
    tokenized_sent: List[str], word: str
) -> List[str]:
    sentence_list = []
    count = 0

    for sentence in tokenized_sent:
        if count == 3:
            break

        match = regex.search(r"\b" + word + r"\b", sentence, flags=regex.IGNORECASE)

        if match is not None:
            sentence_list.append(sentence)
            count += 1

    return sentence_list


def main():
    nltk_setup()
    test_doc_dir = Path.cwd() / "test_docs"
    stopwords_file = Path.cwd() / "stopwords.txt"

    with open(stopwords_file, encoding="utf-8", mode="r") as stopword_file:
        stopwords: List[str] | None = stopword_file.read().splitlines()

    if not stopwords:
        raise FileNotFoundError(f"{stopwords_file} may be missing or empty.")

    text_files_paths: List[Path] = get_files_to_analyze(test_doc_dir)

    # contains the table that will eventually be printed
    # format[( (word, freq), filename, [sentences] ),]
    final_word_list: List[tuple] = []

    for fp in text_files_paths:
        with open(fp, encoding="utf-8", mode="r") as doc:
            doc_str = doc.read()
            tokenized_sentences = st(doc_str)

        # remove stop words from the tokenized list
        tk_list_without_stopwords: List[str] = get_tokenized_list(doc_str, stopwords)

        # Get the frequency of each word
        word_fd = FreqDist(tk_list_without_stopwords)

        top_five_words: List[Tuple[Any, int]] = word_fd.most_common(5)

        final_word_list.extend(
            generate_final_word_list(top_five_words, fp.name, tokenized_sentences)
        )

    create_excel_workbook(final_word_list)


def nltk_setup():
    """
    Set the environment variable for NLTK to the current working directory and
    download the necessary models if not already installed.
    """
    punkt_model_path: Path = Path.cwd() / "tokenizers" / "punkt"
    if not os.environ.get("NLTK_DATA"):
        os.environ["NLTK_DATA"] = str(Path.cwd())

    if not punkt_model_path.exists():
        import nltk

        nltk.download("punkt", str(Path.cwd()))


if __name__ == "__main__":
    main()
