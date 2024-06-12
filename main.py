import fitz
import re
from bs4 import BeautifulSoup
import os
from os import walk
from pandas import DataFrame, ExcelWriter


def display_match(match):
    if match is None:
        return None
    return '<Match: %r, groups=%r>' % (match.group(), match.groups())


def get_pages(document, a, b):
    text = ""
    for i in range(a - 1, b):
        page = document.load_page(i)
        pageText = page.get_text("xhtml")
        text = text + pageText

    return text


def remove_extras(text):
    image = re.compile(r"<p><img.*?</p>", flags=re.DOTALL)
    page = re.compile(r"<p><b>Page.*?</b></p>", flags=re.DOTALL)
    div = re.compile(r"</?div.*?>")

    text = image.sub("", text)
    text = page.sub("", text)
    text = div.sub("", text)

    return text


def remove_extras_html(text):
    image = re.compile(r"<img.*?>", flags=re.DOTALL)
    page = re.compile(r"<p.*?<b>.*?>Page \d+ of \d+.*?</b></p>")
    div = re.compile(r"</?div.*?>")

    text = image.sub("", text)
    text = page.sub("", text)
    text = div.sub("", text)

    return text


def extract_news(text):
    newsBlock = re.compile(r"<p>.*?</p>.*?<p>.*?<b>.*?<p>.*?Document \S+</p>", flags=re.DOTALL)
    news = newsBlock.findall(text)

    return news


def extract_features(newsBlock):

    lines = re.split(r"\n+", newsBlock)

    headingRe = re.compile(r"<b>.*?</b>")
    wordsRe = re.compile(r"\d+[\d|,]* words")
    dateRe = re.compile(r"\d+[\d|,]* words.*?([1-3]?\d (?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|June?|July?|Aug("
                        r"?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?) \d{4})")
    abbreviationRe = re.compile(r"Document ([A-Z]+)")

    try:
        abbreviation = abbreviationRe.findall(lines[-1])[0]
    except IndexError:
        print(lines)

    if headingRe.search(lines[1]):
        heading = lines[1]
    else:
        heading = None

    if wordsRe.search(lines[2]):
        try:
            words = wordsRe.findall(lines[2])[0]
            date = dateRe.findall(lines[2])[0]
            sourceRe = re.compile(rf"{words}.*{date} (?:\d\d:\d\d GMT)?(.*) {abbreviation}")
            source = re.findall(sourceRe, string=lines[2])[0]
            text = lines[3:-1]
            entireBlock = lines[1:]
        except IndexError:

            try:
                words = wordsRe.findall(lines[2])[0]

                dateRe = re.compile(
                    r"([1-3]?\d (?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|June?|July?|Aug("
                    r"?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?) \d{4}).*?\d+[\d|,]* words.*?")
                date = dateRe.findall(lines[2])[0]

                sourceRe = re.compile(rf"{words}\W+?By (.*?)\(English\)")
                source = sourceRe.findall(string=lines[2])[0]
                textTemp = "".join(lines[2])
                textRe = re.compile(rf"\(English\)(.*) Document \w+</p>", flags=re.DOTALL)
                text = textRe.findall(textTemp)[0]
                entireBlock = lines[1:-1]

            except IndexError:

                print("Error in ", lines[-1])
                raise

    else:
        words = None
        date = None
        source = None
        text = None
        entireBlock = None

    featuresList = [heading, words, date, source, text, entireBlock]

    return featuresList


def process_data(dataPath):

    f = []
    for (dirpath, dirnames, filenames) in walk(dataPath):
        f.extend(filenames)
        break

    result = []
    for i in f:
        print(f"processing File {i} ...")
        # Load Document
        path = dataPath + '/' + i
        doc = fitz.open(path)
        startingPage = doc.get_toc()[0][2]
        pageCount = doc.page_count

        # Get Document Pages
        text = get_pages(doc, startingPage, pageCount)

        # Preprocess Document
        text = remove_extras(text)
        newsBlocks = extract_news(text)

        # Process news
        for (index, j) in enumerate(newsBlocks):
            features = extract_features(j)

            heading = BeautifulSoup(features[0], features="html.parser").get_text()
            wordsCount = BeautifulSoup(features[1], features="html.parser").get_text()
            date = BeautifulSoup(features[2], features="html.parser").get_text()
            source = BeautifulSoup(features[3], features="html.parser").get_text()
            text = BeautifulSoup("".join(features[4]), features="html.parser").get_text("\n")
            entireBlock = BeautifulSoup("".join(features[5]), features="html.parser").get_text("\n")

            row = [heading, wordsCount, date, source, text, i]

            result.append(row)

            textFileData = entireBlock
            filename = i + "_" + str(index)
            write_to_text("Output", textFileData, filename)

    return result


def write_to_file(outputPath, fileData):
    df = DataFrame(fileData, columns=["Heading", "Words", "Date", "Source", "Text", "FileName"])
    writer = ExcelWriter(f'{outputPath}/output.xlsx', engine='xlsxwriter')
    df.to_excel(writer, sheet_name='news data', index=False)
    writer.save()
    return


def write_to_text(outputPath, fileData, fileName):

    fileName = re.sub('.pdf', '', fileName)
    with open(outputPath + "/" + f'{fileName}.txt', 'w', encoding="utf-8") as f:
        f.write(fileData)


res = process_data("Data")
write_to_file("Output", res)
