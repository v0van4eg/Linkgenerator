from .megamarket_generator import MegamarketGenerator
from .yandexmarket_generator import YandexmarketGenerator


class GeneratorFactory:
    """Фабрика для создания генераторов документов"""
    @staticmethod
    def create_generator(client_name):
        generators = {
            'Мегамаркет': MegamarketGenerator,
            'ЯндексМаркет': YandexmarketGenerator,
        }

        generator_class = generators.get(client_name)
        if generator_class:
            return generator_class()
        else:
            # Возвращаем генератор по умолчанию
            return MegamarketGenerator()

