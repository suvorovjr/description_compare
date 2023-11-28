import re


def cleanhtml(re_clean: str, html_string: str) -> str:
    """
    Функция для очищения строки от html разметки
    :param re_clean: Регулярное выражение для очищения
    :param html_string: строка с html разметкой
    :return: очищенная строка
    """
    cleantext = re.sub(re_clean, '', html_string)
    spaces_clean = '\n'.join(el.strip() for el in cleantext.split('\n') if el.strip())
    return spaces_clean


def desc_identificate(description: str, compare_description: str) -> float:
    """
    Сравнивает два товара по описанию и высчитывает процент их совпадения
    :param description: описание товара № 1
    :param compare_description: описание товара № 2
    :return: процент совпадения
    """
    description_list = description.split(" ")
    compare_description_list = compare_description.split(" ")
    ident_count = 0
    for word in description_list:
        if word in compare_description_list:
            ident_count += 1
    ident_precent = (ident_count / len(description_list)) * 100
    return ident_precent
