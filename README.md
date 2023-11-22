# Scientific-Article-Extractor
A simple program that extracts the 'Introduction', 'Methods', and 'Results' section from a scientific article. 

![Screenshot (45)](https://github.com/HarithAl-Obaidi/Scientific-Article-Extractor/assets/151098129/c3ea590a-1a5f-4164-9ede-ac3587b32c1d)
## General Information
This program is built for users who need to do sectioned text extraction from a large amount of scientific articles. It generates an excel file with each of the "Introduction", "Methods", and "Results" sections in their own column with a rough copy of the text from the articles. We say rough because the PdfReader used sometimes does not have 100% accuracy with the text extraction and may have some faults with the text. Nevertheless the program will cut down on the time it takes to extract each section from each article greatly.
## Technology
Project is created with:
* python 3.11.5
* PyPDF2 2.11.1
* xlsxwriter 3.1.9
## Installation
The installation for this program is very simple. All that is needed is to download the 'article_extraction.py' file, the aforementioned packages, and run it in python 3.11.5.
## Usage
Again, the usage of this program is also fairly simple.
* Once running, the program will ask the user to input a path to the desired file (hopefully a pdf), a path to the directory the user wants to save the excel file in, and the name of the generated excel file. After that the program should run and the file should be in the specified directory.
* The file generated will not have any modifications done to it. The text wrap and cell size modification appearing in the sample image were done post-program run.
