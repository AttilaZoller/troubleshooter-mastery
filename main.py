from bs4 import BeautifulSoup
import re
from fuzzywuzzy import process
import openpyxl

def main():
    korean_dictionary, english_dictionary = collect_dictionary()

    with open("guide.html",encoding="utf-8",mode="r") as f:
        html = f.read()

    soup = BeautifulSoup(html,'html.parser')
    section_contents = soup.select("div.subSection.detailBox")

    category = {"mastery_category_name":"","mastery_set_name":"","mastery_set_components":""}

    try:
        workbook = openpyxl.load_workbook("Mastery.xlsx", data_only=True)
        worksheet = workbook.active
    except FileNotFoundError:
        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        worksheet.append(["카테고리 분류","세트명","특성 1","특성 2","특성 3","특성 4"])

    for section in section_contents:
        sub_section_title = str(section.select("div.subSectionTitle")[0].text).strip()

        if "Mastery Set" in sub_section_title or "AC.MS" in sub_section_title:
            category["mastery_category_name"] = re.sub(r"(Mastery Set - |AC\.MS - | Class | Type |\(.*\))","",sub_section_title)

            mastery_set_name_contents = section.select("div.subSectionDesc > div")
            mastery_contents = section.select("ul")
            for index, mastery in enumerate(mastery_contents):
                try:
                    category["mastery_set_name"] = mastery_set_name_contents[index].text.replace("-","").strip()
                    category["mastery_set_components"] = mastery.select("li > b")[0].text.strip().split(" + ")
                except IndexError:
                    category["mastery_set_name"] = mastery_set_name_contents[index].text.strip().replace("- ","")
                    category["mastery_set_components"] = section.select("div.subSectionDesc > b")[0].text.strip().split(" + ")
                finally:
                    category["mastery_set_components"] = [mastery.strip() for mastery in category["mastery_set_components"]]
                    english_data = [category["mastery_category_name"],category["mastery_set_name"]] + category["mastery_set_components"]
                    korean_data = translate_english_to_korean(english_data,korean_dictionary,english_dictionary)
                    worksheet.append(korean_data)
                    workbook.save("Mastery.xlsx")


def collect_dictionary():
    with open("dic_keyword.dic",encoding="utf-8",mode="r") as f:
            raw_text = f.readlines()

    korean_dictionary = []
    english_dictionary = []
    for line in raw_text:
        if "!Mastery" in line or "!Job" in line or "!AbilitySubType" in line:
            tag_name = re.compile(r"(?<=\!).*?(?=\])").search(line).group()

            mastery_tags = re.compile(r"{}]".format(tag_name)).findall(line)
            if len(mastery_tags) == 2:
                korean_name = re.compile(r"(?<={}]).*?(?=\[\!{}])".format(tag_name,tag_name)).search(line).group().strip()
                parsed_text = line.replace("{}]{}".format(tag_name,korean_name),"")
                english_name = re.compile(r"(?<={}]).*".format(tag_name)).search(parsed_text).group().strip()

            elif len(mastery_tags) == 1:
                korean_name = re.compile(r"(?<={}]).*?(?=[a-zA-Z])".format(tag_name)).search(line).group().strip()
                parsed_text = line.replace("{}".format(korean_name),"")
                english_name = re.compile(r"(?<={}]).*".format(tag_name)).search(parsed_text).group().strip()

            korean_dictionary.append(korean_name)
            english_dictionary.append(english_name)

    return korean_dictionary, english_dictionary

def translate_english_to_korean(english_data,korean_dictionary,english_dictionary):
    korean_data = []
    for data in english_data:
        highest_match = process.extractOne(data,english_dictionary)
        translated_word = korean_dictionary[english_dictionary.index(highest_match[0])]

        if data == "General":
            korean_data.append("공용")

        elif highest_match[1] == 100:
            korean_data.append(translated_word)

        elif highest_match[1] > 90:
            korean_data.append(translated_word)

        elif highest_match[1] <= 90:
            with open("translating_log.txt",mode="a","utf-8") as f:
                print("{}는 {}({})와(과) 일치율 {}(으)로 영어명을 그대로 사용함.".format(data,highest_match[0],translated_word,str(highest_match[1])))
            korean_data.append(data)

    return korean_data

main()