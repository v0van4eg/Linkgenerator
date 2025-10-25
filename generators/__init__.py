from .megamarket_generator import MegamarketGenerator
from .elise_generator import EliseGenerator
from .yandexmarket_generator import YandexmarketGenerator
from .magnitcosmetic_generator import MagnitcosmeticGenerator


class GeneratorFactory:
    """Фабрика для создания генераторов документов"""

    @staticmethod
    def create_generator(client_name):
        generators = {
            'Мегамаркет': MegamarketGenerator,
            'ЭЛИЗЕ': EliseGenerator,
            'ЯндексМаркет': YandexmarketGenerator,
            'МагнитКосметик': MagnitcosmeticGenerator
        }

        generator_class = generators.get(client_name)
        if generator_class:
            return generator_class()
        else:
            # Возвращаем генератор по умолчанию
            return EliseGenerator()
