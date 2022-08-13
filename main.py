#!/usr/bin/python3

from docx import Document

document = Document()

from modules.load_file import loadfile

def main():
    m = loadfile('text.txt')
    print(m)

main()