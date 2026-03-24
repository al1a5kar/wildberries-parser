import pandas as pd
import json
import time
import re
from typing import List, Dict
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
import warnings
warnings.filterwarnings('ignore')

class WildberriesDetailedParser:
    def __init__(self, headless: bool = False):
        """Инициализация парсера"""
        chrome_options = Options()
        
        if headless:
            chrome_options.add_argument('--headless')
            chrome_options.add_argument('--disable-gpu')
        
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--disable-blink-features=AutomationControlled')
        chrome_options.add_argument('--window-size=1920,1080')
        chrome_options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        self.driver = webdriver.Chrome(options=chrome_options)
        self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        self.wait = WebDriverWait(self.driver, 20)
        
    def search_products(self, query: str, limit: int = 100) -> List[Dict]:
        """Парсинг поисковой выдачи с переходом в каждую карточку"""
        all_products = []
        page = 1
        
        search_url = f"https://www.wildberries.ru/catalog/0/search.aspx?search={query}"
        
        try:
            print(f"🔍 Открываю страницу поиска...")
            self.driver.get(search_url)
            time.sleep(5)
            
            # Прокручиваем для загрузки
            self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(3)
            self.driver.execute_script("window.scrollTo(0, 0);")
            time.sleep(2)
            
            while len(all_products) < limit:
                print(f"\n📄 Парсинг страницы {page}...")
                
                # Находим все карточки товаров
                product_links = self.get_product_links()
                
                if not product_links:
                    print("❌ Товары не найдены")
                    break
                
                print(f"Найдено ссылок на странице: {len(product_links)}")
                
                # Проходим по каждой ссылке
                for idx, link in enumerate(product_links):
                    if len(all_products) >= limit:
                        break
                    
                    try:
                        print(f"  Обработка товара {idx + 1}/{len(product_links)}...")
                        product_data = self.parse_product_page(link)
                        if product_data:
                            all_products.append(product_data)
                            print(f"  ✅ Собрано товаров: {len(all_products)}")
                            print(f"     Страна: {product_data.get('country', 'Не определена')}")
                    except Exception as e:
                        print(f"  ⚠️ Ошибка при обработке: {e}")
                        continue
                
                # Переход на следующую страницу
                if not self.go_to_next_page():
                    print("📌 Достигнут конец выдачи")
                    break
                
                page += 1
                time.sleep(3)
                
        except Exception as e:
            print(f"❌ Ошибка: {e}")
        
        return all_products[:limit]
    
    def get_product_links(self) -> List[str]:
        """Получение ссылок на товары со страницы поиска"""
        links = []
        
        selectors = [
            "a.product-card__link",
            "a[class*='product-card']",
            "a[href*='/catalog/']",
            "div[data-nm-id] a",
            "article a"
        ]
        
        for selector in selectors:
            try:
                elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                for elem in elements:
                    href = elem.get_attribute("href")
                    if href and "/catalog/" in href and "detail.aspx" in href:
                        if href not in links:
                            links.append(href)
                if links:
                    print(f"✅ Найдено {len(links)} ссылок")
                    return links
            except:
                continue
        
        return links
    
    def parse_product_page(self, url: str) -> Dict:
        """Парсинг детальной страницы товара"""
        try:
            print(f"    Открываю карточку товара...")
            self.driver.get(url)
            time.sleep(4)  # Увеличил задержку для полной загрузки
            
            # Ждем загрузки основных данных
            try:
                self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "h1, .product-page__header")))
            except:
                pass
            
            # Собираем данные
            product_data = {}
            
            # Артикул
            product_data['article'] = self.get_article()
            
            # Название
            product_data['name'] = self.get_product_name()
            
            # Цена
            product_data['price'] = self.get_price()
            
            # Описание
            product_data['description'] = self.get_description()
            
            # Изображения
            product_data['images'] = self.get_images()
            
            # Характеристики
            product_data['characteristics'] = self.get_characteristics()
            
            # Продавец
            product_data['seller_name'], product_data['seller_url'] = self.get_seller_info()
            
            # Размеры и остатки
            product_data['sizes'], product_data['stocks'] = self.get_sizes_and_stocks()
            
            # Рейтинг и отзывы
            product_data['rating'], product_data['reviews_count'] = self.get_rating_and_reviews()
            
            # Ссылка на товар
            product_data['url'] = url
            
            # Страна производства - улучшенное определение
            product_data['country'] = self.get_country_improved(
                product_data['name'], 
                product_data['description'],
                product_data['characteristics']
            )
            
            return product_data
            
        except Exception as e:
            print(f"    Ошибка при парсинге страницы: {e}")
            return None
    
    def get_article(self) -> str:
        """Получение артикула"""
        try:
            selectors = [
                "span[class*='article']",
                "div[class*='article']",
                ".product-page__article",
                "[data-article]"
            ]
            
            for selector in selectors:
                try:
                    elem = self.driver.find_element(By.CSS_SELECTOR, selector)
                    text = elem.text
                    match = re.search(r'\d+', text)
                    if match:
                        return match.group()
                except:
                    continue
            
            url = self.driver.current_url
            match = re.search(r'/catalog/(\d+)/', url)
            if match:
                return match.group()
        except:
            pass
        return ""
    
    def get_product_name(self) -> str:
        """Получение названия товара"""
        try:
            selectors = [
                "h1",
                ".product-page__header",
                "[class*='product-name']",
                "[class*='goods-name']"
            ]
            
            for selector in selectors:
                try:
                    elem = self.driver.find_element(By.CSS_SELECTOR, selector)
                    name = elem.text.strip()
                    if name:
                        return name
                except:
                    continue
        except:
            pass
        return ""
    
    def get_price(self) -> float:
        """Получение цены"""
        try:
            selectors = [
                "ins.price__lower-price",
                ".final-price",
                "[class*='price'] ins",
                "span[class*='price']"
            ]
            
            for selector in selectors:
                try:
                    elem = self.driver.find_element(By.CSS_SELECTOR, selector)
                    price_text = elem.text.replace("₽", "").replace(" ", "").replace(" ", "").strip()
                    if price_text and price_text.replace('.', '').isdigit():
                        return float(price_text)
                except:
                    continue
        except:
            pass
        return 0.0
    
    def get_description(self) -> str:
        """Получение описания"""
        try:
            selectors = [
                ".product-description__text",
                "[class*='description']",
                ".j-description"
            ]
            
            for selector in selectors:
                try:
                    elem = self.driver.find_element(By.CSS_SELECTOR, selector)
                    desc = elem.text.strip()
                    if desc:
                        return desc
                except:
                    continue
        except:
            pass
        return ""
    
    def get_images(self) -> str:
        """Получение ссылок на изображения"""
        images = []
        try:
            selectors = [
                "img[class*='photo']",
                ".product-page__image img",
                "[data-image] img"
            ]
            
            for selector in selectors:
                try:
                    imgs = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    for img in imgs[:5]:
                        src = img.get_attribute("src")
                        if src and src.startswith("http"):
                            images.append(src)
                    if images:
                        break
                except:
                    continue
            
            if not images:
                article = self.get_article()
                if article:
                    vol = int(article) // 100000
                    part = int(article) // 1000
                    for i in range(1, 6):
                        img_url = f"https://basket-01.wb.ru/vol{vol}/part{part}/{article}/images/big/{i}.jpg"
                        images.append(img_url)
        except:
            pass
        
        return ', '.join(images)
    
    def get_characteristics(self) -> str:
        """Получение характеристик в JSON формате"""
        chars = {}
        try:
            selectors = [
                ".product-params__item",
                "[class*='characteristics'] li",
                "table.product-details"
            ]
            
            for selector in selectors:
                try:
                    items = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    for item in items:
                        try:
                            name_elem = item.find_element(By.CSS_SELECTOR, "[class*='name'], dt")
                            value_elem = item.find_element(By.CSS_SELECTOR, "[class*='value'], dd")
                            name = name_elem.text.strip()
                            value = value_elem.text.strip()
                            if name and value:
                                chars[name] = value
                        except:
                            continue
                    if chars:
                        break
                except:
                    continue
        except:
            pass
        
        return json.dumps(chars, ensure_ascii=False)
    
    def get_seller_info(self) -> tuple:
        """Получение информации о продавце"""
        seller_name = ""
        seller_url = ""
        
        try:
            selectors = [
                "a[class*='seller']",
                ".seller-info__name",
                "[class*='supplier'] a"
            ]
            
            for selector in selectors:
                try:
                    elem = self.driver.find_element(By.CSS_SELECTOR, selector)
                    seller_name = elem.text.strip()
                    seller_url = elem.get_attribute("href")
                    if seller_url and not seller_url.startswith("http"):
                        seller_url = "https://www.wildberries.ru" + seller_url
                    if seller_name:
                        break
                except:
                    continue
        except:
            pass
        
        return seller_name, seller_url
    
    def get_sizes_and_stocks(self) -> tuple:
        """Получение размеров и остатков"""
        sizes = []
        total_stock = 0
        
        try:
            size_selectors = [
                ".size-list",
                "[class*='sizes']",
                ".j-size-list"
            ]
            
            for selector in size_selectors:
                try:
                    size_elements = self.driver.find_elements(By.CSS_SELECTOR, f"{selector} button, {selector} .size")
                    for elem in size_elements:
                        size_text = elem.text.strip()
                        if size_text and size_text not in sizes:
                            sizes.append(size_text)
                            total_stock += 1
                    if sizes:
                        break
                except:
                    continue
        except:
            pass
        
        return ', '.join(sizes), total_stock
    
    def get_rating_and_reviews(self) -> tuple:
        """Получение рейтинга и количества отзывов"""
        rating = 0.0
        reviews = 0
        
        try:
            rating_selectors = [
                ".product-page__rating span",
                "[class*='rating'] span:first-child",
                ".j-rating"
            ]
            
            for selector in rating_selectors:
                try:
                    elem = self.driver.find_element(By.CSS_SELECTOR, selector)
                    rating_text = elem.text.strip()
                    if rating_text:
                        rating = float(rating_text.replace(',', '.'))
                        break
                except:
                    continue
            
            review_selectors = [
                "a[href*='reviews'] span",
                ".product-page__reviews",
                "[class*='feedback'] span"
            ]
            
            for selector in review_selectors:
                try:
                    elem = self.driver.find_element(By.CSS_SELECTOR, selector)
                    review_text = elem.text.strip()
                    numbers = re.findall(r'\d+', review_text)
                    if numbers:
                        reviews = int(numbers[0])
                        break
                except:
                    continue
                    
        except:
            pass
        
        return rating, reviews
    
    def get_country_improved(self, name: str, description: str, characteristics_json: str) -> str:
        """Улучшенное определение страны производства"""
        # Объединяем все тексты для поиска
        search_text = (name + " " + description).lower()
        
        # Добавляем характеристики
        try:
            chars = json.loads(characteristics_json)
            for key, value in chars.items():
                if key.lower() in ['страна', 'страна производства', 'страна-производитель', 'made in', 'country']:
                    # Если есть явное поле с названием страны
                    if 'россия' in value.lower() or 'russia' in value.lower():
                        return 'Россия'
                    elif 'китай' in value.lower() or 'china' in value.lower():
                        return 'Китай'
                    elif 'италия' in value.lower() or 'italy' in value.lower():
                        return 'Италия'
                    elif 'турция' in value.lower() or 'turkey' in value.lower():
                        return 'Турция'
                    elif 'беларусь' in value.lower() or 'belarus' in value.lower():
                        return 'Беларусь'
                    else:
                        return value  # Возвращаем как есть
                search_text += " " + value.lower()
        except:
            pass
        
        # Расширенный поиск по ключевым словам
        country_patterns = {
            'Россия': [
                'россия', 'russia', 'made in russia', 'сделано в россии', 'российское', 
                'российский', 'рф', 'rf', 'russian federation'
            ],
            'Китай': [
                'китай', 'china', 'made in china', 'сделано в китае', 'китайское'
            ],
            'Италия': [
                'италия', 'italy', 'made in italy', 'сделано в италии', 'итальянское'
            ],
            'Турция': [
                'турция', 'turkey', 'made in turkey', 'сделано в турции', 'турецкое'
            ],
            'Беларусь': [
                'беларусь', 'belarus', 'made in belarus', 'сделано в беларуси', 'белорусское'
            ],
        }
        
        for country, keywords in country_patterns.items():
            for keyword in keywords:
                if keyword in search_text:
                    return country
        
        return 'Не указана'
    
    def go_to_next_page(self) -> bool:
        """Переход на следующую страницу"""
        try:
            next_selectors = [
                "a.pagination-next",
                "[data-page='next']",
                "a[aria-label='Следующая страница']"
            ]
            
            for selector in next_selectors:
                try:
                    next_button = self.driver.find_element(By.CSS_SELECTOR, selector)
                    if next_button and next_button.is_enabled():
                        self.driver.execute_script("arguments[0].click();", next_button)
                        time.sleep(3)
                        return True
                except:
                    continue
        except:
            pass
        return False
    
    def save_to_excel(self, products: List[Dict], filename: str):
        """Сохранение в Excel"""
        if not products:
            print(f"❌ Нет данных для сохранения")
            return
        
        df = pd.DataFrame(products)
        
        column_mapping = {
            'url': 'Ссылка на товар',
            'article': 'Артикул',
            'name': 'Название',
            'price': 'Цена',
            'description': 'Описание',
            'images': 'Ссылки на изображения',
            'characteristics': 'Характеристики',
            'seller_name': 'Название селлера',
            'seller_url': 'Ссылка на селлера',
            'sizes': 'Размеры товара',
            'stocks': 'Остатки по товару',
            'rating': 'Рейтинг',
            'reviews_count': 'Количество отзывов',
            'country': 'Страна производства'
        }
        
        df = df.rename(columns=column_mapping)
        
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Каталог', index=False)
            
            worksheet = writer.sheets['Каталог']
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
        
        print(f"✅ Файл {filename} сохранен ({len(products)} товаров)")
    
    def filter_products(self, products: List[Dict]) -> List[Dict]:
        """Фильтрация товаров по критериям"""
        filtered = []
        
        for product in products:
            try:
                rating = float(product.get('rating', 0))
                price = float(product.get('price', float('inf')))
                country = product.get('country', '')
                
                # Проверяем условия фильтрации
                if rating >= 4.5 and price <= 10000 and country == 'Россия':
                    filtered.append(product)
                    print(f"  ✅ Отфильтрован: {product.get('name', '')[:50]} (рейтинг: {rating}, цена: {price}, страна: {country})")
            except Exception as e:
                print(f"  ⚠️ Ошибка фильтрации: {e}")
                continue
        
        return filtered
    
    def close(self):
        """Закрытие браузера"""
        if self.driver:
            self.driver.quit()
            print("🔒 Браузер закрыт")

def main():
    print("="*80)
    print("Wildberries Детальный парсер - Пальто из натуральной шерсти")
    print("="*80)
    
    parser = None
    
    try:
        print("\n🚀 Запускаю браузер...")
        parser = WildberriesDetailedParser(headless=False)
        
        print("\n🔍 Начинаю сбор данных...")
        print("⏳ Это может занять несколько минут...")
        
        products = parser.search_products("пальто из натуральной шерсти", limit=50)  # Уменьшил лимит для теста
        
        print(f"\n📊 Всего собрано товаров: {len(products)}")
        
        if products:
            # Сохраняем полный каталог
            full_filename = "wildberries_palto_full.xlsx"
            parser.save_to_excel(products, full_filename)
            print(f"\n📁 Полный каталог сохранен: {full_filename}")
            
            # Фильтруем товары
            print(f"\n🎯 Начинаю фильтрацию товаров...")
            print(f"   Критерии: рейтинг >= 4.5, цена <= 10000, страна = Россия")
            
            filtered_products = parser.filter_products(products)
            print(f"\n📊 Результат фильтрации: {len(filtered_products)} товаров")
            
            if filtered_products:
                filtered_filename = "wildberries_palto_filtered.xlsx"
                parser.save_to_excel(filtered_products, filtered_filename)
                print(f"\n✅ Отфильтрованный каталог: {filtered_filename}")
                
                # Выводим статистику по отфильтрованным
                print(f"\n📈 Статистика отфильтрованных товаров:")
                print(f"   - Количество: {len(filtered_products)}")
                if filtered_products:
                    avg_rating = sum(p['rating'] for p in filtered_products) / len(filtered_products)
                    avg_price = sum(p['price'] for p in filtered_products) / len(filtered_products)
                    print(f"   - Средний рейтинг: {avg_rating:.2f}")
                    print(f"   - Средняя цена: {avg_price:.2f} ₽")
            else:
                print("\n⚠️ Нет товаров, соответствующих условиям фильтрации")
                print("   Возможные причины:")
                print("   - Нет товаров из России")
                print("   - Нет товаров с рейтингом >= 4.5")
                print("   - Нет товаров с ценой <= 10000 ₽")
        else:
            print("\n❌ Товары не найдены")
    
    except Exception as e:
        print(f"\n❌ Ошибка: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        if parser:
            parser.close()
    
    print("\n" + "="*80)

if __name__ == "__main__":
    main()