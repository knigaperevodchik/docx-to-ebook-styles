# docx-to-ebook-styles

[RU](#русский) | [EN](#english)

<a name="русский"></a>

Скрипт для подготовки `.docx` файлов к конвертации в **FB2 / EPUB** через [Calibre](https://calibre-ebook.com/).

Просто положи скрипт рядом с документом и запусти — он сам разберётся.

---

## Зачем это нужно

Когда конвертируешь `.docx` в FB2 или EPUB через Calibre, важно чтобы документ был правильно размечен: заголовки через стили `Heading 1/2/3`, абзацы без ручных Tab-отступов и пустых строк-разделителей, метаданные на русском языке. Иначе Calibre не построит оглавление, главы не будут начинаться с новой страницы, а текст будет выглядеть криво на ридере.

Этот скрипт автоматически приводит документ в порядок.

---

## Что делает скрипт

- Настраивает стили `Normal`, `Heading 1/2/3`, `Title`, `Subtitle`, `Quote`, `Footnote Text`
- Добавляет `page-break-before` в стиль `Heading 1` — каждая глава начинается с новой страницы
- Убирает прямое форматирование абзацев (пространство до/после, ручные отступы)
- Удаляет пустые строки-разделители между абзацами (они ломают FB2)
- Убирает Tab в начале абзацев — красная строка задаётся через стиль
- Устанавливает язык документа `ru-RU`
- **Оригинал не трогает** — результат сохраняется как `имя_файла_calibre.docx`

### Параметры стилей

| Элемент | Шрифт | Размер | Форматирование |
|---|---|---|---|
| Основной текст | Times New Roman | 12 pt | По ширине, красная строка 1.25 см, интервал 1.5 |
| Heading 1 (глава) | Times New Roman | 16 pt | Жирный, по центру, **разрыв страницы перед** |
| Heading 2 (раздел) | Times New Roman | 14 pt | Жирный, по центру |
| Heading 3 (подраздел) | Times New Roman | 13 pt | Жирный курсив |
| Цитата | Times New Roman | 12 pt | Курсив, отступ 2 см с обеих сторон |
| Сноска | Times New Roman | 10 pt | По ширине |

---

## Установка

Нужен Python 3 и библиотека `python-docx`:

```bash
pip install python-docx
```

---

## Использование

1. Скачай `применить_стили.py`
2. Положи его в одну папку с твоим `.docx` файлом
3. Запусти:

```bash
python3 применить_стили.py
```

Если в папке несколько `.docx` файлов — скрипт спросит какой обработать, либо можно выбрать «обработать все».

**Результат:** рядом появится файл `имя_файла_calibre.docx` — его и кидай в Calibre.

---

## Следующий шаг: конвертация в Calibre

1. Открой Calibre
2. Добавь `имя_файла_calibre.docx`
3. Нажми **Конвертировать книги**
4. Выбери формат: `FB2` или `EPUB`
5. Calibre автоматически подхватит заголовки `H1/H2/H3` для оглавления

---

## Настройки

В начале файла можно поменять параметры под себя:

```python
BODY_FONT    = "Times New Roman"  # шрифт основного текста
BODY_SIZE_PT = 12                 # размер шрифта
LINE_SPACING = 1.5                # межстрочный интервал
FIRST_LINE   = 1.25               # красная строка в см
```

---

## Поддержать автора

Если скрипт оказался полезным — буду рад поддержке 🙏

| Способ | Реквизиты |
|---|---|
| ☕ Boosty | [boosty.to/knigaperevodchik](https://boosty.to/knigaperevodchik) |
| 💎 TON / USDT (TON) | `UQBWKwf2mgakNi4Ls2I6NNs1okcDyCxivdxxc22ypsMV4590` |
| 💵 USDT (TRC-20) | `TDdok5FgB6fJSXZrPzxnn7hMk4qREUZPJe` |

---

## Лицензия

MIT — делай что хочешь, упоминание автора приветствуется.

---

<a name="english"></a>

# docx-to-ebook-styles

A script that prepares `.docx` files for conversion to **FB2 / EPUB** via [Calibre](https://calibre-ebook.com/).

Just drop the script next to your document and run it — it figures out the rest.

---

## Why you need this

When converting `.docx` to FB2 or EPUB via Calibre, the document needs proper markup: headings using `Heading 1/2/3` styles, paragraphs without manual Tab indents or blank separator lines, and metadata in the correct language. Without this, Calibre can't build a table of contents, chapters won't start on new pages, and the text will look broken on an e-reader.

This script fixes all of that automatically.

---

## What the script does

- Configures styles: `Normal`, `Heading 1/2/3`, `Title`, `Subtitle`, `Quote`, `Footnote Text`
- Adds `page-break-before` to `Heading 1` — every chapter starts on a new page in Calibre
- Removes direct paragraph formatting (manual spacing, indents)
- Deletes blank separator lines between paragraphs (they break FB2 output)
- Removes Tab characters at the start of paragraphs — first-line indent is set via style
- Sets document language to `ru-RU`
- **Does not modify the original** — result is saved as `filename_calibre.docx`

### Style parameters

| Element | Font | Size | Formatting |
|---|---|---|---|
| Body text | Times New Roman | 12 pt | Justified, first-line indent 1.25 cm, 1.5 line spacing |
| Heading 1 (chapter) | Times New Roman | 16 pt | Bold, centered, **page break before** |
| Heading 2 (section) | Times New Roman | 14 pt | Bold, centered |
| Heading 3 (subsection) | Times New Roman | 13 pt | Bold italic |
| Quote | Times New Roman | 12 pt | Italic, 2 cm indent on both sides |
| Footnote | Times New Roman | 10 pt | Justified |

---

## Installation

Requires Python 3 and `python-docx`:

```bash
pip install python-docx
```

---

## Usage

1. Download `применить_стили.py`
2. Place it in the same folder as your `.docx` file
3. Run:

```bash
python3 применить_стили.py
```

If there are multiple `.docx` files in the folder, the script will ask which one to process, or you can choose to process all of them.

**Result:** a new file `filename_calibre.docx` will appear next to the original — use that one in Calibre.

---

## Next step: converting in Calibre

1. Open Calibre
2. Add `filename_calibre.docx`
3. Click **Convert books**
4. Choose output format: `FB2` or `EPUB`
5. Calibre will automatically pick up `H1/H2/H3` headings to build the table of contents

---

## Configuration

Edit the constants at the top of the file to suit your needs:

```python
BODY_FONT    = "Times New Roman"  # body text font
BODY_SIZE_PT = 12                 # font size in points
LINE_SPACING = 1.5                # line spacing multiplier
FIRST_LINE   = 1.25               # first-line indent in cm
```

---

## Support the author

If this script was useful, donations are welcome 🙏

| Method | Details |
|---|---|
| ☕ Boosty | [boosty.to/knigaperevodchik](https://boosty.to/knigaperevodchik) |
| 💎 TON / USDT (TON) | `UQBWKwf2mgakNi4Ls2I6NNs1okcDyCxivdxxc22ypsMV4590` |
| 💵 USDT (TRC-20) | `TDdok5FgB6fJSXZrPzxnn7hMk4qREUZPJe` |

---

## License

MIT — do whatever you want, attribution is appreciated.
