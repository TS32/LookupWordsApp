# Lookup English words in the text files and generate word document contains the explanations/sentences examples #

娃的英语作业中有一项是完成单词表查词，每次的词语很多，如果一个个查，费时费力。所以写了这个脚本，从脚本所在的目录下的input目录中，读取txt文件中要查的词（txt文件中每行一个单词或者短语）。然后调用金山词霸的查词接口，获取查词结果和例句。结果以表格的方式放在output目录下的word文档里。

## Purpose ##

This script file reads all the txt files in the "input" folder, then try to get all the English words in those files, generate a wordlist, lookup every word from the internet , get the pos and acceptations, example sentences and genrate a word document for each text file contains all these information. it's effecient if you have a lot of words to translate.

## Requirements ##

In order to run this scripts, Python3.6+ is needed. and use the requirements.txt to install the requested libraries, following the commands:

pip3 install -r requirements.txt


(note: it's ony tested in win10+python3.8)

## How to run ##

```Console logs

LookupWordsApp> & C:/Python38/python.exe "c:/python/LookupWordsApp/LookupWordsApp.py"

[Info]: -Run on Windows-

[Info]: Current script absolute filename c:\python\LookupWordsApp\LookupWordsApp.py

[Info]: Current script absolute root path c:\python\LookupWordsApp

[Info]: Input folder path c:\python\LookupWordsApp\input

[Info]: Output folder path c:\python\LookupWordsApp\output


fileLoader found 1 files (totalFileSize=1 KB) with pattern [.txt] in folder c:\python\LookupWordsApp\input

Start 1 files jobs from input folder c:\python\LookupWordsApp\input


Delete c:\python\LookupWordsApp\output directory!


Create c:\python\LookupWordsApp\output directory!


          [Info]: file c:\python\LookupWordsApp\input\Top 3000 US English Words.txt encoding type is ascii


          [Info]: 152 words loaded!

Lookup words: 100%|██████████████████████████████████████████████████████████████████| 152/152 [00:56<00:00,  2.69words/s] 

Vocabulary translation document c:\python\LookupWordsApp\output\Top 3000 US English Words.docx generated successfully!


Done! 1 files processed successfully! 

```

Look at the docx file in 'output' folder to understand the result format. 