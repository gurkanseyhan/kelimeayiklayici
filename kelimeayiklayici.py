# coding=utf-8
import docx
from collections import Counter


replace_list = ['.', ',', '?',
                '!', '"', '(',
                ')', "'", '-',
                '_', ':', ';'
                ]


def get_key_piece(filename):
    document = ""
    tem_document = docx.Document(filename)
    for _par in tem_document.paragraphs:
        document += _par.text
    clear_punctuation(document)


def clear_punctuation(doc_text):
    for _rep in replace_list:
        doc_text = doc_text.replace(_rep, " ")
    return clear_repeat(doc_text.split(" "))


def clear_repeat(punc_clear_list):
    a = Counter(punc_clear_list)
    for _aa in a.keys():
        if _aa.strip() != "":
            print(_aa + " : " + str(a[_aa]))

    print ("Farklı Kelime Sayısı : " + str(len(a)))


get_key_piece("ornekmetin.docx")
