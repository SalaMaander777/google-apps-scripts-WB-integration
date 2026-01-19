# Статистика карточек товаров (Sales Funnel) - API Documentation

## Endpoint
```
POST https://seller-analytics-api.wildberries.ru/api/analytics/v3/sales-funnel/products
```

## Описание
Метод формирует отчёт о товарах, сравнивая ключевые показатели — например, добавления в корзину, заказы и переходы в карточку товара — за текущий период с аналогичным прошлым.

## Лимиты запросов
- **1 минута**: 3 запроса
- **Интервал**: 20 секунд
- **Всплеск**: 3 запроса

## Request

### Headers
```
Authorization: <WB_API_TOKEN>
Content-Type: application/json
```

### Request Body (JSON)

```json
{
  "selectedPeriod": {
    "start": "2024-01-18",
    "end": "2024-01-18"
  },
  "pastPeriod": {
    "start": "2023-01-18",
    "end": "2023-01-18"
  },
  "nmIds": [],
  "brandNames": [],
  "subjectIds": [],
  "tagIds": [],
  "skipDeletedNm": false,
  "orderBy": {
    "field": "openCard",
    "mode": "desc"
  },
  "limit": 1000,
  "offset": 0
}
```

### Параметры

#### selectedPeriod (required)
- **Тип**: object
- **Описание**: Запрашиваемый период
  - `start` (required, string date): Начало периода
  - `end` (required, string date): Конец периода

#### pastPeriod (optional)
- **Тип**: object
- **Описание**: Период для сравнения
  - `start` (required, string date): Начало периода
  - `end` (required, string date): Конец периода
  - **Примечание**: Данные в pastPeriod указаны за такой же период, что и в selectedPeriod
  - Если дата начала pastPeriod раньше, чем год назад от текущей даты, она будет приведена к виду: `pastPeriod.start = текущая дата — 365 дней`

#### nmIds (optional)
- **Тип**: Array of integers (0-1000 items)
- **Описание**: Артикулы WB, по которым нужно составить отчёт. Оставьте пустым `[]`, чтобы получить отчёт обо всех товарах

#### brandNames (optional)
- **Тип**: Array of strings
- **Описание**: Список брендов для фильтрации

#### subjectIds (optional)
- **Тип**: Array of integers
- **Описание**: Список ID предметов для фильтрации

#### tagIds (optional)
- **Тип**: Array of integers
- **Описание**: Список ID ярлыков для фильтрации

#### skipDeletedNm (optional)
- **Тип**: boolean
- **Описание**: Скрыть удалённые карточки товаров

#### orderBy (optional)
- **Тип**: object
- **Описание**: Параметры сортировки
  - `field` (required, string): Поле для сортировки
  - `mode` (required, string): Порядок сортировки (`asc` или `desc`)

**Доступные поля для сортировки:**
- `openCard` — Перешли в карточку
- `addToCart` — Положили в корзину
- `orderCount` — Заказали товаров, шт
- `orderSum` — Заказали на сумму
- `buyoutCount` — Выкупили товаров, шт
- `buyoutSum` — Выкупили на сумму
- `cancelCount` — Отменили товаров, шт
- `cancelSum` — Отменили на сумму
- `avgPrice` — Средняя цена
- `stockMpQty` — Остатки на складах продавца, шт
- `stockWbQty` — Остатки на складах WB, шт
- `shareOrderPercent` — Доля в выручке
- `addToWishlist` — Добавили в Отложенные
- `timeToReady` — Среднее время доставки
- `localizationPercent` — Локальные заказы в рамках одного региона
- `wbClub.orderCount` — Заказали товаров с WB Клубом, шт
- `wbClub.orderSum` — Заказали на сумму с WB Клубом
- `wbClub.buyoutSum` — Выкупили товаров с WB Клубом, шт
- `wbClub.buyoutCount` — Процент выкупа с WB Клубом
- `wbClub.cancelSum` — Отменили товаров с WB Клубом, шт
- `wbClub.avgPrice` — Средняя цена с WB Клубом
- `wbClub.buyoutPercent` — Процент выкупа с WB Клубом
- `wbClub.avgOrderCountPerDay` — Среднее количество заказов в день с WB Клубом, шт
- `wbClub.cancelCount` — Отменили товаров с WB Клубом, шт

#### limit (optional)
- **Тип**: integer
- **Default**: 50
- **Max**: 1000
- **Описание**: Количество карточек товара в ответе

#### offset (optional)
- **Тип**: integer
- **Default**: 0
- **Описание**: Сколько элементов пропустить (для пагинации)

### Особенности фильтрации
- Параметры `brandNames`, `subjectIds`, `tagIds`, `nmIds` могут быть пустыми `[]`, тогда в ответе возвращаются все карточки продавца
- Если вы указали несколько параметров, в ответе будут карточки, в которых есть одновременно все эти параметры
- Если карточки не подходят по параметрам запроса, вернётся пустой ответ `[]`
- Можно получить отчёт максимум за последние 365 дней

## Response

### Успешный ответ (200 OK)

```json
{
  "data": {
    "products": [
      {
        "product": {
          "nmId": 15756350,
          "title": "Пальто зимнее женское пуховик",
          "vendorCode": "DWC0029/темно-серый",
          "brandName": "Cloud Concept",
          "subjectId": 170,
          "subjectName": "Пальто",
          "tags": [],
          "productRating": 8.7,
          "feedbackRating": 0,
          "stocks": {
            "wb": 0,
            "mp": 0,
            "balanceSum": 0
          }
        },
        "statistic": {
          "selected": {
            "period": {
              "start": "2026-01-17",
              "end": "2026-01-17"
            },
            "openCount": 0,
            "cartCount": 0,
            "orderCount": 0,
            "orderSum": 0,
            "buyoutCount": 0,
            "buyoutSum": 0,
            "cancelCount": 0,
            "cancelSum": 0,
            "avgPrice": 0,
            "avgOrdersCountPerDay": 0,
            "shareOrderPercent": 0,
            "addToWishlist": 0,
            "timeToReady": {
              "days": 0,
              "hours": 0,
              "mins": 0
            },
            "localizationPercent": 0,
            "wbClub": {
              "orderCount": 0,
              "orderSum": 0,
              "buyoutSum": 0,
              "buyoutCount": 0,
              "cancelSum": 0,
              "cancelCount": 0,
              "avgPrice": 0,
              "buyoutPercent": 0,
              "avgOrderCountPerDay": 0
            },
            "conversions": {
              "addToCartPercent": 0,
              "cartToOrderPercent": 0,
              "buyoutPercent": 0
            }
          },
          "past": {
            "period": {
              "start": "2025-01-18",
              "end": "2025-01-18"
            },
            "openCount": 1,
            "cartCount": 0,
            "orderCount": 0,
            "orderSum": 0,
            "buyoutCount": 0,
            "buyoutSum": 0,
            "cancelCount": 0,
            "cancelSum": 0,
            "avgPrice": 0,
            "avgOrdersCountPerDay": 0,
            "shareOrderPercent": 0,
            "addToWishlist": 0,
            "timeToReady": {
              "days": 0,
              "hours": 0,
              "mins": 0
            },
            "localizationPercent": 0,
            "wbClub": {
              "orderCount": 0,
              "orderSum": 0,
              "buyoutSum": 0,
              "buyoutCount": 0,
              "cancelSum": 0,
              "cancelCount": 0,
              "avgPrice": 0,
              "buyoutPercent": 0,
              "avgOrderCountPerDay": 0
            },
            "conversions": {
              "addToCartPercent": 0,
              "cartToOrderPercent": 0,
              "buyoutPercent": 0
            }
          },
          "comparison": {
            "openCountDynamic": -100,
            "cartCountDynamic": 0,
            "orderCountDynamic": 0,
            "orderSumDynamic": 0,
            "buyoutCountDynamic": 0,
            "buyoutSumDynamic": 0,
            "cancelCountDynamic": 0,
            "cancelSumDynamic": 0,
            "avgOrdersCountPerDayDynamic": 0,
            "avgPriceDynamic": 0,
            "shareOrderPercentDynamic": 0,
            "addToWishlistDynamic": 0,
            "timeToReadyDynamic": {
              "days": 0,
              "hours": 0,
              "mins": 0
            },
            "localizationPercentDynamic": 0,
            "wbClubDynamic": {
              "orderCount": 0,
              "orderSum": 0,
              "buyoutSum": 0,
              "buyoutCount": 0,
              "cancelSum": 0,
              "cancelCount": 0,
              "avgPrice": 0,
              "buyoutPercent": 0,
              "avgOrderCountPerDay": 0
            },
            "conversions": {
              "addToCartPercent": 0,
              "cartToOrderPercent": 0,
              "buyoutPercent": 0
            }
          }
        }
      }
    ]
  }
}
```

### Структура данных

#### data (required, object)
Корневой объект ответа

#### data.products (required, Array)
Массив товаров со статистикой

#### product (required, object)
Информация о карточке товара
- `nmId` (integer): Артикул WB
- `title` (string): Название карточки товара
- `vendorCode` (string): Артикул продавца
- `brandName` (string): Бренд
- `subjectId` (integer): ID предмета
- `subjectName` (string): Название предмета
- `tags` (array): Ярлыки
- `productRating` (number): Оценка карточки
- `feedbackRating` (number): Оценка пользователей
- `stocks` (object): Остатки
  - `wb` (integer): Остатки на складах WB, шт
  - `mp` (integer): Остатки на складах МП, шт
  - `balanceSum` (integer): Сумма остатков на складах, руб

#### statistic (required, object)
Статистика по товару

#### statistic.selected (required, object)
Статистика за запрашиваемый период
- `period` (object): Период
  - `start` (string): Дата начала
  - `end` (string): Дата конца
- `openCount` (integer): Количество переходов в карточку товара
- `cartCount` (integer): Положили в корзину, шт
- `orderCount` (integer): Заказали товаров, шт
- `orderSum` (integer): Заказали на сумму
- `buyoutCount` (integer): Выкупили товаров, шт
- `buyoutSum` (integer): Выкупили на сумму
- `cancelCount` (integer): Отменили товаров, шт
- `cancelSum` (integer): Отменили на сумму
- `avgPrice` (integer): Средняя цена
- `avgOrdersCountPerDay` (number): Среднее количество заказов в день, шт
- `shareOrderPercent` (number): Доля в выручке
- `addToWishlist` (integer): Добавили в Отложенные
- `timeToReady` (object): Среднее время доставки
  - `days` (integer): Дни
  - `hours` (integer): Часы
  - `mins` (integer): Минуты
- `localizationPercent` (integer): Локальные заказы в рамках одного региона
- `wbClub` (object): Статистика WB Клуба
  - `orderCount` (integer): Заказали товаров с WB Клубом, шт
  - `orderSum` (integer): Заказали на сумму с WB Клубом
  - `buyoutSum` (integer): Выкупили на сумму с WB Клубом
  - `buyoutCount` (integer): Выкупили товаров с WB Клубом, шт
  - `cancelSum` (integer): Отменили на сумму с WB Клубом
  - `cancelCount` (integer): Отменили товаров с WB Клубом, шт
  - `avgPrice` (integer): Средняя цена с WB Клубом
  - `buyoutPercent` (integer): Процент выкупа с WB Клубом
  - `avgOrderCountPerDay` (number): Среднее количество заказов в день с WB Клубом, шт
- `conversions` (object): Конверсии
  - `addToCartPercent` (integer): Конверсия в корзину, %
  - `cartToOrderPercent` (integer): Конверсия в заказ, %
  - `buyoutPercent` (integer): Процент выкупа, %

#### statistic.past (object)
Статистика за период для сравнения (аналогичная структура как у selected)

#### statistic.comparison (object)
Сравнение selected с past
- `openCountDynamic` (number): Динамика переходов в карточку, %
- `cartCountDynamic` (number): Динамика добавлений в корзину, %
- `orderCountDynamic` (number): Динамика заказов, %
- `orderSumDynamic` (number): Динамика суммы заказов, %
- `buyoutCountDynamic` (number): Динамика выкупов, %
- `buyoutSumDynamic` (number): Динамика суммы выкупов, %
- `cancelCountDynamic` (number): Динамика отмен, %
- `cancelSumDynamic` (number): Динамика суммы отмен, %
- `avgOrdersCountPerDayDynamic` (number): Динамика среднего количества заказов в день, %
- `avgPriceDynamic` (number): Динамика средней цены, %
- `shareOrderPercentDynamic` (number): Динамика доли в выручке, %
- `addToWishlistDynamic` (number): Динамика добавлений в отложенные, %
- `timeToReadyDynamic` (object): Динамика времени доставки
- `localizationPercentDynamic` (number): Динамика локальных заказов, %
- `wbClubDynamic` (object): Динамика WB Клуба
- `conversions` (object): Динамика конверсий

## Примеры использования

### Получить все товары за вчерашний день
```javascript
var yesterday = getPreviousDay(); // "2024-01-18"
var yearAgo = getDateYearAgo(yesterday); // "2023-01-18"

var payload = {
  selectedPeriod: {
    start: yesterday,
    end: yesterday
  },
  pastPeriod: {
    start: yearAgo,
    end: yearAgo
  },
  nmIds: [],
  brandNames: [],
  subjectIds: [],
  tagIds: [],
  skipDeletedNm: false,
  orderBy: {
    field: "openCard",
    mode: "desc"
  },
  limit: 1000,
  offset: 0
};
```

### Получить товары конкретного бренда
```javascript
var payload = {
  selectedPeriod: {
    start: "2024-01-18",
    end: "2024-01-18"
  },
  pastPeriod: {
    start: "2023-01-18",
    end: "2023-01-18"
  },
  brandNames: ["MyBrand"],
  limit: 1000,
  offset: 0
};
```

### Пагинация
```javascript
// Первая страница
var payload1 = {
  selectedPeriod: { start: "2024-01-18", end: "2024-01-18" },
  limit: 1000,
  offset: 0
};

// Вторая страница
var payload2 = {
  selectedPeriod: { start: "2024-01-18", end: "2024-01-18" },
  limit: 1000,
  offset: 1000
};
```

## Обработка ошибок

### 204 No Content
Нет данных за указанный период

### 400 Bad Request
Неверные параметры запроса

### 401 Unauthorized
Неверный токен авторизации

### 429 Too Many Requests
Превышен лимит запросов

## Примечания
- Максимальный период выгрузки: 365 дней от текущей даты
- Период для сравнения автоматически корректируется, если выходит за пределы 365 дней
- Для получения всех товаров используйте пагинацию с limit=1000
- API поддерживает только POST запросы
