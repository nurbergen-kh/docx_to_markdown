# DOCX to Markdown Converter

Набор скриптов для конвертации документов DOCX в формат Markdown.

## Описание

Этот проект содержит несколько скриптов для извлечения текста из DOCX файлов и преобразования его в формат Markdown. Каждый скрипт использует разные библиотеки и подходы для обработки документов.

## Скрипты

### 1. docs_to_markdown.py

Основной скрипт для конвертации DOCX в Markdown с использованием библиотеки `python-docx`. Этот скрипт обрабатывает текст, таблицы и форматирование.

```bash
poetry run python docs_to_markdown.py input.docx output.md
```

### 2. docx_to_html_markdown.py

Скрипт, который сначала конвертирует DOCX в HTML с помощью библиотеки `mammoth`, а затем преобразует HTML в Markdown с помощью `html2text`. Этот подход лучше сохраняет форматирование.

```bash
poetry run python docx_to_html_markdown.py input.docx output_html_md.md
```

### 3. extract_text_docx2python.py

Скрипт для извлечения текста из DOCX файлов с использованием библиотеки `docx2python`. Этот скрипт фокусируется на извлечении чистого текста.

```bash
poetry run python extract_text_docx2python.py input.docx
```

### 4. extract_text_mammoth.py

Скрипт для извлечения текста из DOCX файлов с использованием библиотеки `mammoth`. Этот скрипт также фокусируется на извлечении чистого текста.

```bash
poetry run python extract_text_mammoth.py input.docx
```

## Установка

Проект использует Poetry для управления зависимостями. Для установки выполните:

```bash
# Установка Poetry (если еще не установлен)
curl -sSL https://install.python-poetry.org | python3 -

# Установка зависимостей
poetry install
```

## Зависимости

- python-docx
- docx2python
- mammoth
- html2text
- beautifulsoup4
- tabulate

## Особенности

- Обработка таблиц с сохранением структуры
- Форматирование заголовков
- Обработка ссылок
- Сохранение форматирования текста

## Ограничения

- Сложное форматирование может быть потеряно при конвертации
- Некоторые специальные элементы документа могут не обрабатываться корректно
- Изображения извлекаются как альтернативный текст

## Лицензия

MIT 