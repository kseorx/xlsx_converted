"""
Конфигурация для загрузки Excel-отчётов в ClickHouse.
Настройте параметры подключения и пути перед запуском.
"""

# ==================== ClickHouse ====================
CH_HOST = "localhost"       # Хост ClickHouse
CH_PORT = 9000              # Порт (9000 — native, 8123 — HTTP)
CH_USER = "default"         # Пользователь
CH_PASSWORD = ""            # Пароль
CH_DATABASE = "default"     # База данных
CH_SECURE = False           # True для SSL-соединения

# ==================== Пути ====================
DATA_DIR = "/Users/xeroks/Projects/convert/data"  # Папка с Excel-файлами

# ==================== Маппинг названий форм ====================
# Ключи — подстроки, которые ищем в названии листа Excel
# Значение — каноническое имя формы для КХД
FORM_PATTERNS = {
    "И.1":    "И1",
    "И.2":    "И2",
    "И.34":   "И34",
    "И.35":   "И35",
    "КОНС.1": "КОНС1",
}

# Необязательные переопределения для файлов без года/типа в имени.
# Пример: "special.xlsx": {"organization": "Организация", "year": 2025, "type": "план"}
FILE_METADATA = {}

# ==================== Схема таблицы ClickHouse ====================
# Одна таблица для всех форм
CREATE_TABLE_SQL = """
CREATE TABLE IF NOT EXISTS budget_reports
(
    id                  Int64,
    pp_number           Nullable(String),
    article_code        Nullable(String),
    article_name        String,
    ekbs_code           Nullable(String),
    metric_unit         Nullable(String),
    value_year_minus2   Nullable(Float64),
    value_year_minus1   Nullable(Float64),
    value_current_year  Nullable(Float64),
    value_jan           Nullable(Float64),
    value_feb           Nullable(Float64),
    value_mar           Nullable(Float64),
    value_q1            Nullable(Float64),
    value_apr           Nullable(Float64),
    value_may           Nullable(Float64),
    value_jun           Nullable(Float64),
    value_q2            Nullable(Float64),
    value_jul           Nullable(Float64),
    value_aug           Nullable(Float64),
    value_sep           Nullable(Float64),
    value_q3            Nullable(Float64),
    value_oct           Nullable(Float64),
    value_nov           Nullable(Float64),
    value_dec           Nullable(Float64),
    value_q4            Nullable(Float64),
    value_6m            Nullable(Float64),
    value_9m            Nullable(Float64),
    date_year           String,
    form_type           String,
    organization_name   String,
    info_type           String,
    source_id           String,
    version_id          String
)
ENGINE = MergeTree()
ORDER BY (date_year, form_type, organization_name, info_type, id)
SETTINGS index_granularity = 8192
"""

# ==================== Маппинг заголовков колонок Excel ====================
# Для каждой формы — список ожидаемых заголовков в порядке следования
# Позиции начинаются с 0 (после метаданных: № п/п, Код статьи, ...)

# И.1 / И.34 / КОНС.1 — формат: prev_year, curr_year, Q1-Q4, 6мес, 9мес
FORMATTER_Q = [
    "prev_year_fact", "curr_year_val",
    "q1", "q2", "q3", "q4", "half_year", "nine_months"
]

# И.2 / И.35 — формат: budget_total, fact_total, variance,
#   budget_jan, fact_jan, var_jan, ... budget_dec, fact_dec, var_dec
FORMATTER_M = [
    "budget_total", "fact_total", "variance_total",
    "budget_jan", "fact_jan", "variance_jan",
    "budget_feb", "fact_feb", "variance_feb",
    "budget_mar", "fact_mar", "variance_mar",
    "budget_apr", "fact_apr", "variance_apr",
    "budget_may", "fact_may", "variance_may",
    "budget_jun", "fact_jun", "variance_jun",
    "budget_jul", "fact_jul", "variance_jul",
    "budget_aug", "fact_aug", "variance_aug",
    "budget_sep", "fact_sep", "variance_sep",
    "budget_oct", "fact_oct", "variance_oct",
    "budget_nov", "fact_nov", "variance_nov",
    "budget_dec", "fact_dec", "variance_dec",
]

# ==================== Поведение ====================
BATCH_SIZE = 10000          # Строк за один INSERT
DRY_RUN = False             # True = только показать, что будет загружено
