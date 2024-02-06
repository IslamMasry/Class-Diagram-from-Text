import spacy
import os
import pyodbc
import pandas as pd

def connect_to_database():
    return pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ="classattr.accdb;')

def retrieve_class_names(cursor, software_id):
    cursor.execute(f'SELECT class_name FROM classes WHERE software_id = {software_id}')
    return [row[0] for row in cursor.fetchall()]

def process_file(cursor, file_path, class_names, nlp, output_directory):
    software_id = int(os.path.basename(file_path).replace(".txt", ""))
    class_names_in_file = retrieve_class_names(cursor, software_id)

    with open(file_path, encoding="utf8", errors='ignore') as file_handle:
        lines = file_handle.read()
        doc = nlp(lines)

        feature_vector = []
        for sent in doc.sents:
            for token in nlp(sent.text):
                if token.is_alpha and not token.is_stop:
                    feature_vector.append(create_feature_vector(token, class_names_in_file, lines))

        for i in range(len(feature_vector) - 1):
            if feature_vector[i + 1][3] == 'NOUN':
                feature_vector[i][12] = 1

        df = pd.DataFrame(feature_vector)
        df.to_csv(os.path.join(output_directory, f"{software_id}.csv"), index=False, header=False)

def create_feature_vector(token, class_names, lines):
    is_oov = token.is_oov
    word = token.text
    vector = token.vector_norm
    lemma = token.lemma_
    pos = token.pos_
    tag = token.tag_
    dependency = token.dep_
    shape = token.shape_
    is_alpha = token.is_alpha
    is_stop = token.is_stop
    is_noun = 1 if token.pos_ == "NOUN" else 0
    has_s = 1 if "'" in token.text else 0
    term_freq = lines.count(token.text)
    label = "class" if token.text in class_names and is_noun == 1 else "none"

    return [word, vector, lemma, pos, tag, dependency, shape, is_alpha, is_stop, is_noun, has_s, term_freq, 0, is_oov, label]

def main():
    conn = connect_to_database()
    cursor = conn.cursor()

    nlp = spacy.load("en_core_web_lg")
    files_directory = r'CSVFiles'
    output_directory = r'CSVFiles'
    for software_id in range(1, 70):
        class_names = retrieve_class_names(cursor, software_id)
        file_name = f"{software_id}.txt"
        file_path = os.path.join(files_directory, file_name)
        process_file(cursor, file_path, class_names, nlp, output_directory)

    conn.close()

if __name__ == "__main__":

    main()
