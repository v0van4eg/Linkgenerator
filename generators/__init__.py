# generators/__init__.py (пример обновленной фабрики)
from .megamarket_generator import MegamarketGenerator
from .yandexmarket_generator import YandexmarketGenerator
# Импортируйте другие генераторы по мере необходимости

class GeneratorFactory:
    @staticmethod
    def create_generator(template_name): # Изменено: client_name -> template_name
        """Создает экземпляр генератора на основе имени шаблона"""
        generators = {
            'Мегамаркет': MegamarketGenerator,
            'ЯндексМаркет': YandexmarketGenerator,
            # Добавьте другие шаблоны и соответствующие классы
        }

        generator_class = generators.get(template_name) # Изменено: client_name -> template_name
        if generator_class:
            return generator_class()
        else:
            # Возвращаем базовый генератор или вызываем ошибку, если шаблон не найден
            raise ValueError(f"Неизвестный шаблон: {template_name}") # Изменено: клиента -> шаблона

# Экспортируйте фабрику
__all__ = ['GeneratorFactory']
