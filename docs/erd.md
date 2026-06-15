# ERD · Финансовый аналитик

Ссылка на dbdiagram: (нарисуй на dbdiagram.io и вставь сюда)

## Таблицы

### `users`

| Колонка    | Тип         | PK/FK | Примечание |
|------------|-------------|-------|------------|
| id         | bigint      | PK    |
| tg_id      | bigint      |       | UNIQUE — id пользователя Telegram |
| name       | text        |       | имя из профиля Telegram |
| created_at | timestamptz |       | default now() |

### `statements`

| Колонка | Тип | PK/FK | Примечание |
|---------|-----|-------|------------|
| id | bigint | PK | |
| user_id | bigint | FK → users.id | чья выписка |
| period | text | | например «2026-05» |
| file_name | text | | имя загруженного PDF |
| uploaded_at | timestamptz | | default now() |

### `transactions`

| Колонка | Тип | PK/FK | Примечание |
|---------|-----|-------|------------|
| id | bigint | PK | |
| statement_id | bigint | FK → statements.id | из какой выписки |
| op_date | date | | дата операции |
| amount | numeric | | сумма (минус = расход) |
| description | text | | назначение из выписки |
| category_id | bigint | FK → categories.id | категория (может быть пустой) |

### `categories`

| Колонка | Тип | PK/FK | Примечание |
|---------|-----|-------|------------|
| id | bigint | PK | |
| name | text | | еда, аренда, развлечения, переводы, прочее |
| keywords | text | | слова для авто-определения категории |

## Связи

- `users` 1 — N `statements` через `statements.user_id`
- `statements` 1 — N `transactions` через `transactions.statement_id`
- `categories` 1 — N `transactions` через `transactions.category_id`

## Проверка spec → ERD

- F1 «Загрузка и разбор» → строка в `statements` + строки в `transactions`. ✅
- F2 «Сводка по категориям» → `SELECT name, SUM(amount) FROM transactions JOIN categories GROUP BY name`. ✅

Каждая фича из spec ложится на таблицу — значит, spec и ERD не противоречат.

## SQL

Черновик в `schema/init.sql`. В Supabase сегодня не запускаем — это День 2 с ментором.