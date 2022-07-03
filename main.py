import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape
import pandas as pd
import collections
from os import getcwd, getenv
from dotenv import load_dotenv


YEAR_OF_INCORPORATION = 1920


def sort_wines_by_categories(wines):
    category_dict_out_of_main_dict = collections.defaultdict(list)
    for dictionary in wines:
        category_dict_out_of_main_dict[dictionary.get('Категория')].append(
            dict(list(dictionary.items())[1:])
        )
    return category_dict_out_of_main_dict


def count_age_of_company(year):
    return datetime.datetime.now().year-year


def proper_translation_of_year_into_russian_on_basis_of_company_age(year):
    proper_words_in_russian_dict = ['лет', 'год', 'года']
    exceptions = [11, 12, 13, 14]
    first_special_condition_numbers = [1]
    second_special_condition_numbers = [2, 3, 4]
    last_number_of_age_of_company = int(str(year)[-1])
    if last_number_of_age_of_company in first_special_condition_numbers \
            and year not in exceptions:
        proper_word_in_russian = proper_words_in_russian_dict[1]
    elif last_number_of_age_of_company in second_special_condition_numbers \
            and year not in exceptions:
        proper_word_in_russian = proper_words_in_russian_dict[2]
    else:
        proper_word_in_russian = proper_words_in_russian_dict[0]
    return proper_word_in_russian


def main():
    load_dotenv()
    where_assortment = getenv('PATH_TO_ASSORTMENT_TABLE', getcwd())
    assortment = pd.read_excel(
        f'{where_assortment}/wine.xlsx',
        sheet_name='Лист1',
        keep_default_na=False
    ).to_dict(orient='records')
    categories_of_wines = sort_wines_by_categories(
        assortment
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
        categories_of_wines=categories_of_wines
    )
    with open('index.html', 'w', encoding='utf8') as file:
        file.write(rendered_page)
    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
