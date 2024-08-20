import os
import pandas as pd


def read_lines_from_file(file_path):
    """Reads lines from a file and returns them as a list."""
    with open(file_path, 'r', encoding='utf-8') as file:

        return file.readlines()


def create_dataframe(english_lines, farsi_lines):
    """Creates a DataFrame from English and Farsi lines."""
    return pd.DataFrame({
        'English': english_lines,
        'Farsi': farsi_lines
    })


def save_dataframe_to_csv(dataframe, output_path):
    """Saves the DataFrame to a CSV file."""
    dataframe.to_csv(output_path, index=False)


def process_file_pairs():
    """Processes file pairs and returns a list of generated CSV files."""
    file_pairs = [
        ('NeuLab-TedTalks.en-fa.en', 'NeuLab-TedTalks.en-fa.fa'),
        ('ELRC_2922.en-fa.en', 'ELRC_2922.en-fa.fa'),
        ('GNOME.en-fa.en', 'GNOME.en-fa.fa'),
        ('infopankki.en-fa.en', 'infopankki.en-fa.fa'),
        ('QED.en-fa.en', 'QED.en-fa.fa'),
        ('Tanzil.en-fa.en', 'Tanzil.en-fa.fa'),
        ('TED2020.en-fa.en', 'TED2020.en-fa.fa'),
        ('TEP.en-fa.en', 'TEP.en-fa.fa'),
        ('tico-19.en-fa.en', 'tico-19.en-fa.fa'),
        ('tldr-pages.en-fa.en', 'tldr-pages.en-fa.fa'),
        ('MIZAN.en-fa.en', 'MIZAN.en-fa.fa'),
        ('Ubuntu.en-fa.en', 'Ubuntu.en-fa.fa'),
        ('WikiMatrix.en-fa.en', 'WikiMatrix.en-fa.fa'),
        ('wikimedia.en-fa.en', 'wikimedia.en-fa.fa'),
        ('XLEnt.en-fa.en', 'XLEnt.en-fa.fa'),
        ('OpenSubtitles.en-fa.en', 'OpenSubtitles.en-fa.fa'),
        ('LinguaTools-WikiTitles.en-fa.en', 'LinguaTools-WikiTitles.en-fa.fa'),
    ]

    data_source_dir = 'DataSource'
    output_files = []

    for en_file, fa_file in file_pairs:
        en_file_path = os.path.join(data_source_dir, en_file)
        fa_file_path = os.path.join(data_source_dir, fa_file)
        output_csv_path = os.path.join(data_source_dir, en_file.replace('.en', '') + '.csv')

        en_lines = read_lines_from_file(en_file_path)
        fa_lines = read_lines_from_file(fa_file_path)

        df = create_dataframe(en_lines, fa_lines)
        df['Source'] = en_file.split('.')[0]

        save_dataframe_to_csv(df, output_csv_path)
        output_files.append(output_csv_path)

    return output_files


def process_tsv_file(tsv_file_path):
    """Processes a TSV file and saves it as a CSV."""
    df = pd.read_csv(tsv_file_path, delimiter='\t')
    df['Source'] = os.path.basename(tsv_file_path).split('.')[0]
    df.columns = ['num1', 'Farsi', 'num2', 'English', 'Source']
    df = df[['English', 'Farsi', 'Source']]
    output_csv_path = tsv_file_path.replace('.tsv', '.csv')
    save_dataframe_to_csv(df, output_csv_path)
    return output_csv_path


def main():
    # Process file pairs
    generated_files = process_file_pairs()
    print("Generated CSV files:", generated_files)

    # Process the TSV file
    tsv_file_path = 'DataSource/Sentence pairs in Persian-English - 2024-08-15.tsv'
    tsv_csv_path = process_tsv_file(tsv_file_path)
    print(f"TSV file saved as CSV: {tsv_csv_path}")


if __name__ == "__main__":
    main()
