-- Финансовый аналитик · черновик схемы
-- НЕ запускать в Supabase сегодня — это День 2 с ментором

create table users (
  id          bigint generated always as identity primary key,
  tg_id       bigint unique not null,        -- id пользователя Telegram
  name        text,                          -- имя из профиля
  created_at  timestamptz default now()
);

create table categories (
  id        bigint generated always as identity primary key,
  name      text not null,                   -- еда, аренда, развлечения, переводы, прочее
  keywords  text                             -- слова для авто-определения категории
);

create table statements (
  id           bigint generated always as identity primary key,
  user_id      bigint not null references users(id),
  period       text,                         -- например «2026-05»
  file_name    text,                         -- имя загруженного PDF
  uploaded_at  timestamptz default now()
);

create table transactions (
  id            bigint generated always as identity primary key,
  statement_id  bigint not null references statements(id),
  op_date       date,                        -- дата операции
  amount        numeric,                     -- сумма (минус = расход)
  description   text,                        -- назначение из выписки
  category_id   bigint references categories(id)   -- может быть пустой
);

-- стартовые категории
insert into categories (name, keywords) values
  ('Еда',          'magnum,small,продукты,кафе,ресторан,glovo,wolt'),
  ('Аренда',       'аренда,квартира,ksk,коммунал'),
  ('Развлечения',  'кино,steam,игры,подписка,netflix'),
  ('Переводы',     'перевод,kaspi gold,p2p'),
  ('Прочее',       '');