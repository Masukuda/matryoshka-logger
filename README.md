# 🪆 Backtest Auto-Logger

Логгер сделок из **TradingView** для **Excel** и **Google Таблицы**, разработанный специально для ведения журналов.

Скрипт считывает параметры инструмента **«Длинная / Короткая позиция»** и индикаторов Матрёшка, определяет торговую сессию и исход сделки по свечам и некоторые другие параметры.

---

## 🚀 Установка за 1 минуту

### Шаг 1. Установите менеджер скриптов Tampermonkey

Выберите и установите расширение **Tampermonkey** для вашего браузера:

* 🌐 [Google Chrome / Brave / Яндекс.Браузер / Opera](https://chromewebstore.google.com/detail/tampermonkey/dhdgffkkebhmkfjojejmpbldmpobfkfo)
* 🦊 [Mozilla Firefox](https://addons.mozilla.org/ru/firefox/addon/tampermonkey/)
* 🧭 [Microsoft Edge](https://microsoftedge.microsoft.com/addons/detail/tampermonkey/iikmkjmpaadaobahmlepeloendndfphd)

> ⚠️ **Обязательная настройка для Chrome, Edge, Brave, Opera и Яндекс.Браузера:**
> В современных браузерах на базе Chromium пользовательские скрипты по умолчанию заблокированы до включения двух переключателей:
> 1. Откройте в браузере страницу **`Расширения`** ➔ **`Управление расширениями`** (или введите в адресную строку `chrome://extensions/`).
> 2. В правом верхнем углу включите тумблер **«Режим разработчика»** (*Developer mode*).
> 3. На карточке **Tampermonkey** нажмите кнопку **«Сведения»** (*Details*), прокрутите чуть ниже и включите тумблер **«Разрешить пользовательские скрипты»** (*Allow User Scripts*).

### Шаг 2. Установите скрипт в один клик

Нажмите на прямую ссылку ниже — Tampermonkey автоматически откроет окно установки:

👉 [**УСТАНОВИТЬ AUTO-LOGGER (Нажмите сюда)**](https://raw.githubusercontent.com/Masukuda/matryoshka-logger/main/matryoshka.user.js)

В открывшемся окне расширения нажмите кнопку **«Установить» (Install)**.

### Шаг 3. Обновите страницу TradingView

Откройте график на [TradingView](https://www.tradingview.com/) и обновите страницу (`F5`). Скрипт готов к работе!

---

## 📊 Как пользоваться и загружать сделки в Excel / Google Таблицы

### 1. Сбор сделки на графике TradingView
