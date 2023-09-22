import os
import docx2python
import PyPDF2
import pandas as pd
import string
import matplotlib.pyplot as plt  # Import matplotlib



def count_words_chars(filename):
    """Counts the number of words and characters in a PDF or DOC file.

    Args:
        filename: The path to the PDF or DOC file.

    Returns:
        A tuple containing the word count, character count (with/without spaces),
        character count without punctuation, and character count without special characters.
    """

    if filename.endswith(".docx"):
        document = docx2python.read_docx(filename)
        text = document.text
    elif filename.endswith(".pdf"):
        text = read_pdf(filename)
    else:
        raise Exception("Unsupported file format: {}".format(filename))

    # Calculate word count
    words = text.split()
    word_count = len(words)

    # Calculate character count without spaces
    char_count_without_spaces = len(text.replace(" ", ""))

    # Remove punctuation
    text_no_punct = text.translate(str.maketrans('', '', string.punctuation))

    # Calculate character count without punctuation
    char_count_without_punct = len(text_no_punct)

    # Remove special characters (excluding spaces and alphanumeric characters)
    text_no_special = ''.join(c for c in text if c.isalnum() or c.isspace())

    # Calculate character count without special characters
    char_count_without_special = len(text_no_special)

    return word_count, char_count_without_spaces, char_count_without_punct, char_count_without_special


    if filename.endswith(".docx"):
        document = docx2python.read_docx(filename)
        word_count = document.word_count
        character_count = document.character_count
    elif filename.endswith(".pdf"):
        text = read_pdf(filename)
        word_count = len(text.split())
        character_count = len(text)
    else:
        raise Exception("Unsupported file format: {}".format(filename))

    return word_count, character_count

def read_pdf(filename):
    """Reads the contents of a PDF file.

    Args:
        filename: The path to the PDF file.

    Returns:
        A string containing the contents of the PDF file.
    """

    pdf_file = open(filename, "rb")
    pdf_reader = PyPDF2.PdfReader(pdf_file)
    text = ""
    for page in pdf_reader.pages:
        text += page.extract_text()
    return text

def generate_report(results, format):
    """Generates a report of various text statistics for each file in the given results.

    Args:
        results: A list of tuples, where each tuple contains the filename and various text statistics.
        format: The format of the report, either "txt" or "csv".

    Returns:
        A string containing the report, or a Pandas DataFrame if the format is "csv".
    """

    if format == "txt":
        report = ""
        report += "File | Words | Char (No Spaces) | Char (No Punct) | Char (No Special) | Average Word Length | Average Sentence Length | Average Paragraph Length\n"
        report += "-----|-------|-------------------|-----------------|-------------------|---------------------|-----------------------|-------------------------"
        for result in results:
            report += "\n{0} | {1} | {2} | {3} | {4} | {5:.2f} | {6:.2f} | {7:.2f}".format(*result)
        return report

    elif format == "csv":
        df = pd.DataFrame(results, columns=["Filename", "Word Count", "Char (No Spaces)", "Char (No Punct)", "Char (No Special)",
                                            "Average Word Length", "Average Sentence Length", "Average Paragraph Length"])
        return df

    else:
        raise Exception("Unsupported format: {}".format(format))

def main():
    """Counts the word and character count of all PDF and DOC files in a given directory."""

    # Ask the user for the directory to scan.
    directory = input("Enter the path to the directory to scan: ")

    # Check if the directory exists.
    if not os.path.exists(directory):
        print("Directory does not exist.")
        return

    # Create a list to store the results.
    results = []

    # Loop through all of the files in the directory.
    for filename in os.listdir(directory):

        # Check if the file is a PDF or DOC file.
        if filename.endswith(".pdf") or filename.endswith(".docx"):

            # Count the words and characters in the file.
            file_path = os.path.join(directory, filename)
            word_count, char_count_without_spaces, char_count_without_punct, char_count_without_special = count_words_chars(file_path)

            # Calculate average word length
            average_word_length = char_count_without_spaces / word_count if word_count > 0 else 0

            # Calculate average sentence length (assuming each paragraph is a sentence)
            average_sentence_length = word_count

            # Calculate average paragraph length
            average_paragraph_length = 1  # Each paragraph is treated as a sentence

            # Add the results to the list.
            results.append((filename, word_count, char_count_without_spaces, char_count_without_punct, char_count_without_special,
                            average_word_length, average_sentence_length, average_paragraph_length))

    # Ask the user for the report format.
    report_format = input("Enter the report format (txt or csv): ")

    # Generate the report.
    report = generate_report(results, report_format)

    # Ask the user if they want to save the report to a file.
    save_report = input("Do you want to save the report to a file? (y/n) ")
    if save_report.lower() == "y":

        # Get the filename from the user.
        filename = input("Enter the filename: ")

        # Ensure the CSV report has the ".csv" extension
        if report_format == "csv" and not filename.endswith(".csv"):
            filename += ".csv"

        # Save the report to the file.
        if report_format == "txt":
            with open(filename, "w") as f:
                f.write(report)
        elif report_format == "csv":
            report.to_csv(filename, index=False)

    # Print the report to the terminal.
    print(report)

if __name__ == "__main__":
    main()
