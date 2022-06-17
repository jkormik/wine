import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape
import pandas as pd
import collections
from os import getcwd


def make_category_dict_out_of_wine_dict(wines_dict):
    category_dict_out_of_main_dict = collections.defaultdict(list)
    for dictionary in wines_dict:
        category = dictionary.get("Категория")
        del dictionary["Категория"]
        category_dict_out_of_main_dict[category].append(dictionary)
    return category_dict_out_of_main_dict


def sort_category_dict_out_of_wine_dict(category_dict_out_of_wine_dict):
    sorted_category_dict_out_of_wine_dict = collections.defaultdict(list)
    for element in sorted(category_dict_out_of_wine_dict.keys()):
        sorted_category_dict_out_of_wine_dict[element] = \
            category_dict_out_of_wine_dict[element]
    return sorted_category_dict_out_of_wine_dict


def count_age_of_company(year):
    current_year = datetime.datetime.now()
    age_of_company = current_year.year-year.year
    return age_of_company


def proper_translation_of_year_into_russian_on_basis_of_company_age(year):
    proper_words_in_russian_dict = ["лет", "год", "года"]
    exceptions = [11, 12, 13, 14]
    special_condition_numbers_for_god = [1]
    special_condition_numbers_for_goda = [2, 3, 4]
    last_number_of_age_of_company = int(str(year)[-1])

    if last_number_of_age_of_company in special_condition_numbers_for_god \
            and year not in exceptions:
        proper_word_in_russian = proper_words_in_russian_dict[1]
    elif last_number_of_age_of_company in special_condition_numbers_for_goda \
            and year not in exceptions:
        proper_word_in_russian = proper_words_in_russian_dict[2]
    else:
        proper_word_in_russian = proper_words_in_russian_dict[0]

    return proper_word_in_russian


if __name__ == "__main__":
    YEAR_OF_INCORPORATION = datetime.datetime(
        year=1920,
        month=1,
        day=1,
        hour=1
    )
    excel_data_df = pd.read_excel(
        f"{getcwd()}\\wine.xlsx",
        sheet_name="Лист1",
        keep_default_na=False
    )
    excel_data_dict = excel_data_df.to_dict(orient="records")

    category_dict_out_of_wine_dict = make_category_dict_out_of_wine_dict(
        excel_data_dict
    )
    sorted_category_dict_out_of_wine_dict = \
        sort_category_dict_out_of_wine_dict(
            category_dict_out_of_wine_dict
        )
    age_of_company = count_age_of_company(YEAR_OF_INCORPORATION)
    proper_word_in_russian = \
        proper_translation_of_year_into_russian_on_basis_of_company_age(
            age_of_company
        )

    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('template.html')

    rendered_page = template.render(
        age_of_company=age_of_company,
        year_translation=proper_word_in_russian,
        wines_dict=sorted_category_dict_out_of_wine_dict
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()
