#!/usr/bin/python3


def loadfile(filePath):
    with open(filePath, 'r', encoding='utf-8') as f:
        fichier = f.read()
        file = fichier.split('\n')
    return file

if __name__ == '__main__':
    for line in loadfile('../text.txt'):
        print(line)
