import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape
import pandas as pd
import collections
from os import getcwd, getenv
from dotenv import load_dotenv


YEAR_OF_INCORPORATION = 1920


def group_wines_by_categories(wines):
    wines_grouped_by_categories = collections.defaultdict(list)
    for dictionary in wines:
        wine_category = dictionary.pop('Категория')
        
        wines_grouped_by_categories[wine_category].append(
            dictionary
        )
    return wines_grouped_by_categories


def get_proper_rus_year_noun(year):
    rus_year_nouns = ['лет', 'год', 'года']
    exceptions = [11, 12, 13, 14]
    last_year_num = int(str(year)[-1])
    if last_year_num == 1 \
            and year not in exceptions:
        proper_word_in_russian = rus_year_nouns[1]
    elif last_year_num in [2,3,4] \
            and year not in exceptions:
        proper_word_in_russian = rus_year_nouns[2]
    else:
        proper_word_in_russian = rus_year_nouns[0]
    return proper_word_in_russian


def main():
    load_dotenv()
    path_to_assortment = getenv('PATH_TO_ASSORTMENT_TABLE', getcwd())
    assortment = pd.read_excel(
        f'{path_to_assortment}/wine.xlsx',
        sheet_name='Лист1',
        keep_default_na=False
    ).to_dict(orient='records')
    categories_of_wines = group_wines_by_categories(
        assortment
    )
    age_of_company = datetime.datetime.now().year-YEAR_OF_INCORPORATION
    proper_word_in_russian = \
        get_proper_rus_year_noun(
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
