# -*- coding: utf8 -*-
import sys
import docx
from collections import Counter
reload(sys)
sys.setdefaultencoding("utf-8")

alphabet = {
    'A': 'a', 'B': 'b', 'C': 'c',
    'Ç': 'ç', 'D': 'd', 'E': 'e',
    'F': 'f', 'G': 'g', 'Ğ': 'ğ',
    'H': 'h', 'I': 'ı', 'İ': 'i',
    'J': 'j', 'K': 'k', 'L': 'l',
    'M': 'm', 'N': 'n', 'O': 'o',
    'Ö': 'ö', 'P': 'p', 'R': 'r',
    'Ş': 'ş', 'S': 's', 'T': 't',
    'U': 'u', 'Ü': 'ü', 'V': 'v',
    'Y': 'y', 'Z': 'z'
}


replace_list = ['.', ',', '?',
                '!', '"', '(',
                ')', "'", '-',
                '_', ':', ';'
                ]


def get_key_piece(filename):
    document = ""
    tem_document = docx.Document(filename)
    for _par in tem_document.paragraphs:
        document = document + " " + _par.text
    clear_punctuation(document)


def clear_punctuation(doc_text):
    for _test in alphabet.keys():
        doc_text = doc_text.replace(_test, alphabet[_test])

    for _rep in replace_list:
        doc_text = doc_text.replace(_rep, " ")
    return clear_repeat(doc_text.split(" "))


def clear_repeat(punc_clear_list):
    a = Counter(punc_clear_list)
    for _aa in a.keys():
        if _aa.strip() != "" and _aa.strip != " ":
            print(_aa.strip() + " : " + str(a[_aa]))

    print ("Farklı Kelime Sayısı : " + str(len(a)))


filename = raw_input("Lütfen dosya adını uzantısı ile beraber yazınız : ")
get_key_piece(filename)
