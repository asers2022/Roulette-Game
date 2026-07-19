# 📁 Структура проекта Рулетка VBA

Полное описание всех файлов и папок проекта.

---

## 🎯 Обзор структуры

```
Roulette-Game/
│
├── 📄 README.md                    ← Главная документация (ЧИТАЙТЕ ПЕРВЫМ)
├── 📄 QUICKSTART.md                ← Быстрый старт (5 минут)
├── 📄 INSTALLATION.md              ← Подробная инструкция установки
├── 📄 PROJECT_STRUCTURE.md         ← Этот файл (структура проекта)
│
├── 📁 src/                         ← ОСНОВНЫЕ ФАЙЛЫ КОДА
│   ├── RouletteGame.bas            ← Основная логика игры
│   ├── GameInterface.bas           ← Обработчики кнопок
│   └── SpreadsheetLayout.bas       ← Создание интерфейса
│
├── 📁 docs/                        ← ДОКУМЕНТАЦИЯ (опционально)
│   ├── API.md                      ← Описание всех функций
│   ├── STRATEGIES.md               ← Подробно о стратегиях
│   └── TROUBLESHOOTING.md          ← Решение проблем
│
├── 📁 examples/                    ← ПРИМЕРЫ (опционально)
│   ├── Example_Game_Session.xlsx   ← Пример сохраненной игры
│   └── Gameplay_Screenshots.txt    ← Описание скриншотов
│
├── 📄 LICENSE                      ← Лицензия проекта
└── 📄 .gitignore                   ← Git конфиг (игнорируемые файлы)
```

---

## 📋 Описание всех файлов

### 📌 Файлы документации (верхний уровень)

#### **README.md** (Главный файл)
- 📖 Полная документация проекта
- 🎯 Описание всех возможностей
- 📍 Расположение элементов на листе (очень подробно)
- 💰 Типы ставок и выплаты
- 📊 Стратегии игры
- 🎲 Вероятности
- **Рекомендация**: Прочитайте первым!

**Размер**: ~29 кб  
**Время чтения**: 20-30 минут

---

#### **QUICKSTART.md** (Быстрый старт)
- ⚡ Установка за 5 минут
- 🎮 Как начать играть
- 📍 Где находятся элементы
- 💡 Советы для новичков
- 🎯 Примеры игровых сценариев
- 🆘 Быстрое решение проблем
- **Рекомендация**: Для нетерпеливых

**Размер**: ~10 кб  
**Время чтения**: 5-10 минут

---

#### **INSTALLATION.md** (Подробная установка)
- 📖 Пошаговая инструкция с описаниями
- ✅ Системные требования
- 📥 Загрузка файлов
- 🔧 Добавление VBA кода
- 🎮 Первый запуск
- 🆘 Решение проблем
- **Рекомендация**: Для начинающих

**Размер**: ~18 кб  
**Время чтения**: 15-20 минут

---

#### **PROJECT_STRUCTURE.md** (Этот файл)
- 📁 Описание структуры проекта
- 📄 Описание всех файлов
- 🔧 Технические детали
- 🎓 Для разработчиков

**Размер**: ~15 кб  
**Время чтения**: 10-15 минут

---

### 💻 Файлы исходного кода (папка src/)

#### **RouletteGame.bas** (Основная логика - 12.8 кб)

**Содержит:**
- Глобальные переменные и константы
- Структуры данных (GameState, BetRecord)
- Основные функции игры

**Главные функции:**

```vba
' Инициализация
Sub InitializeGame(Optional startingBalance = 1000)
  - Инициализирует игровое состояние
  - Устанавливает начальный баланс
  - Очищает историю

' Вращение колеса
Function SpinWheel() As Integer
  - Генерирует случайное число (0-36)
  - Обновляет счетчик вращений

' Проверка цвета
Function IsRed(number As Integer) As Boolean
Function IsBlack(number As Integer) As Boolean
Function GetColorName(number As Integer) As String
  - Определяют цвет числа

' Основная ставка
Function PlaceBet(betAmount, betType, betValue) As Currency
  - Главная функция игры
  - Вычитает ставку из баланса
  - Вращает колесо
  - Ра��считывает выигрыш
  - Записывает в историю

' Расчет выигрыша
Function CalculateWinnings(spinResult, betType, betValue, betAmount) As Currency
  - Определяет выплату в зависимости от типа ставки
  - Поддерживает 10+ типов ставок

' Запись истории
Sub RecordBet(betAmount, spinResult, betType, betValue, winnings, resultColor)
  - Записывает ставку в историю

' Обновление дисплея
Sub UpdateGameDisplay()
Sub UpdateHistoryDisplay()
  - Обновляет значения на листе

' Статистика
Sub ShowGameStats()
  - Показывает окно с статистикой

' Сброс
Sub ResetGame()
  - Перезагружает игру
```

**Используемые константы:**

```vba
WHEEL_NUMBERS = 37          ' Числа 0-36
RED_NUMBERS = "1,3,5,..."   ' 18 красных чисел
BLACK_NUMBERS = "2,4,6,..." ' 18 черных чисел
```

**Структуры данных:**

```vba
Type GameState
  CurrentBalance As Currency      ' Текущий баланс
  StartingBalance As Currency     ' Начальный баланс
  BetAmount As Currency           ' Размер ставки
  SpinResult As Integer           ' Результат вращения
  LastWinnings As Currency        ' Последний выигрыш
  TotalBets As Currency           ' Всего ставок
  TotalWinnings As Currency       ' Всего выигрышей
  TotalSpins As Integer           ' Количество вращений
  GameActive As Boolean           ' Активна ли игра
  CurrentStrategy As String       ' Текущая стратегия
End Type

Type BetRecord
  BetNumber As Integer            ' Номер ставки
  SpinNumber As Integer           ' Выпавшее число
  BetType As String               ' Тип ставки
  BetValue As String              ' Значение ставки
  BetAmount As Currency           ' Размер ставки
  Winnings As Currency            ' Выигрыш
  ResultColor As String           ' Цвет числа
  BetTime As String               ' Время ставки
End Type
```

---

#### **GameInterface.bas** (Обработчики кнопок - 6.7 кб)

**Содержит:** Все обработчики кликов кнопок

**Группы функций:**

1️⃣ **Основные ставки** (Красное/Черное, Четное/Нечетное)
```vba
Sub BtnRed_Click()              ' Ставка на красное
Sub BtnBlack_Click()            ' Ставка на черное
Sub BtnEven_Click()             ' Ставка на четное
Sub BtnOdd_Click()              ' Ставка на нечетное
```

2️⃣ **High/Low ставки**
```vba
Sub BtnHigh_Click()             ' Ставка на HIGH (19-36)
Sub BtnLow_Click()              ' Ставка на LOW (1-18)
```

3️⃣ **Дюжины**
```vba
Sub BtnDozen1_Click()           ' Дюжина 1 (1-12)
Sub BtnDozen2_Click()           ' Дюжина 2 (13-24)
Sub BtnDozen3_Click()           ' Дюжина 3 (25-36)
```

4️⃣ **Колонны**
```vba
Sub BtnColumn1_Click()          ' Колонна 1
Sub BtnColumn2_Click()          ' Колонна 2
Sub BtnColumn3_Click()          ' Колонна 3
```

5️⃣ **Специальные ставки**
```vba
Sub BtnSingleNumber_Click()     ' На конкретное число
Sub BtnAutoSpin_Click()         ' Автоматические вращения
Sub BtnClearHistory_Click()     ' Очистить историю
```

6️⃣ **Управление игрой**
```vba
Sub BtnStartGame_Click()        ' Начать игру
Sub BtnStatistics_Click()       ' Показать статистику
Sub BtnResetGame_Click()        ' Сброс игры
Sub BtnViewHistory_Click()      ' Открыть историю
Sub BtnExit_Click()             ' Выход
```

7️⃣ **Стратегии**
```vba
Sub BtnStrategyClassic_Click()      ' Классическая
Sub BtnStrategyMartingale_Click()   ' Мартингейл
Sub BtnStrategyDAlembert_Click()    ' Д'Аламбер
```

8️⃣ **Утилиты**
```vba
Function GetBetAmount() As Currency
  - Получает сумму ставки из ячейки D3
  - Валидирует диапазон (1-100)
```

---

#### **SpreadsheetLayout.bas** (Создание интерфейса - 17.2 кб)

**Содержит:** Создание визуального интерфейса игры

**Главная функция:**

```vba
Sub SetupGame()
  - Главный макрос для запуска
  - Вызывает CreateGameLayout()
  - Вызывает InitializeGame(10000)
```

**Функции создания листов:**

```vba
Sub CreateGameLayout()
  - Создает оба листа ("Игра" и "История")
  - Вызывает SetupGameSheet()
  - Вызывает SetupHistorySheet()

Sub SetupGameSheet(ws As Worksheet)
  - Создает все элементы на листе "Игра"
  - Добавляет статистику (левая часть)
  - Добавляет ввод ставки (средняя часть)
  - Создает кнопки (правая часть)
  - Рисует визуализацию колеса

Sub SetupHistorySheet(ws As Worksheet)
  - Создает структуру листа "История"
  - Добавляет заголовки столбцов
  - Форматирует внешний вид
```

**Функция создания кнопок:**

```vba
Sub CreateButton(ws, startCell, endCell, caption, macroName, color, fontSize)
  - Создает одну кнопку
  - Устанавливает цвет и размер шрифта
  - Привязывает макрос к кнопке
  - Параметры:
    ws - лист
    startCell - верхняя левая ячейка (например "F3")
    endCell - нижняя правая ячейка (например "G4")
    caption - текст на кнопке
    macroName - имя макроса для вызова
    color - RGB цвет (например RGB(255,0,0))
    fontSize - размер шрифта
```

**Комментарии в коде:**

В начале файла (~80 строк) подробно описано:
- Где находятся все элементы (по ячейкам)
- Схема расположения
- Таблицы с координатами

---

### 📁 Папка docs/ (Дополнительная документация)

*Опционально - может быть добавлена позже*

#### **docs/API.md** (Если будет создан)
- Описание всех функций и процедур
- Параметры и возвращаемые значения
- Примеры использования

#### **docs/STRATEGIES.md** (Если будет создан)
- Подробное описание всех стратегий
- Математика за стратегиями
- Рекомендации по использованию

#### **docs/TROUBLESHOOTING.md** (Если будет создан)
- Расширенное решение проблем
- FAQ (часто задаваемые вопросы)
- Контакты для поддержки

---

### 📁 Папка examples/ (Примеры)

*Опционально - для демонстрации*

#### **examples/Example_Game_Session.xlsx**
- Пример сохраненной игры
- С полной историей ставок
- Для тестирования

#### **examples/Gameplay_Screenshots.txt**
- Описание скриншотов
- Как выглядит интерфейс
- Примеры результатов

---

### 📄 Другие файлы

#### **LICENSE**
- Лицензия проекта (обычно MIT)
- Условия использования

#### **.gitignore**
- Файлы которые Git игнорирует
- Например: `*.xlsm` (сохраненные файлы игры)

---

## 🔧 Технические детали

### Зависимости

```
✅ Microsoft Excel (встроенный VBA)
✅ Windows/macOS операционная система
❌ Внешние библиотеки НЕ требуются
```

### Размер файлов

| Файл | Размер | Строк кода |
|------|--------|-----------|
| RouletteGame.bas | 12.8 кб | ~350 строк |
| GameInterface.bas | 6.7 кб | ~200 строк |
| SpreadsheetLayout.bas | 17.2 кб | ~500 строк |
| **ВСЕГО** | **36.7 кб** | **~1050 строк** |

### Версионирование

```
Версия 1.0 - 2026-07-19
├── Основная механика рулетки ✅
├── 10+ типов ставок ✅
├── История ставок ✅
├── 3 стратегии ✅
└── Визуальный интерфейс ✅
```

---

## 📊 Диаграмма взаимодействия

```
┌─────────────────────────────────────────────────────────┐
│ ПОЛЬЗОВАТЕЛЬ                                            │
│ (нажимает кнопки на листе Excel)                        │
└─────────────────────────┬───────────────────────────────┘
                          │
                          ▼
┌─────────────────────────────────────────────────────────┐
│ GameInterface.bas                                       │
│ (обработчики кнопок: BtnRed_Click, BtnBlack_Click...)  │
│                                                         │
│ ├─ Получает сумму ставки (GetBetAmount)                │
│ ├─ Вызывает PlaceBet()                                  │
│ └─ Обновляет дисплей                                   │
└─────────────────────────┬───────────────────────────────┘
                          │
                          ▼
┌─────────────────────────────────────────────────────────┐
│ RouletteGame.bas                                        │
│ (основная логика)                                       │
│                                                         │
│ ├─ PlaceBet()                                           │
│ │  ├─ SpinWheel() → генерирует число                   │
│ │  ├─ CalculateWinnings() → считает выигрыш           │
│ │  ├─ RecordBet() → записывает в историю              │
│ │  └─ UpdateGameDisplay() → обновляет лист             │
│ │                                                       │
│ ├─ InitializeGame() → инициализация                    │
│ ├─ ShowGameStats() → статистика                        │
│ └─ LoadStrategy() → выбор стратегии                    │
└─────────────────────────┬───────────────────────────────┘
                          │
                          ▼
┌─────────────────────────────────────────────────────────┐
│ SpreadsheetLayout.bas                                   │
│ (отображение на листе)                                  │
│                                                         │
│ ├─ SetupGame() → создает всё                            │
│ ├─ CreateGameLayout() → оба листа                       │
│ ├─ SetupGameSheet() → лист "Игра"                       │
│ ├─ SetupHistorySheet() → лист "История"               │
│ └─ CreateButton() → кнопки                              │
└─────────────────────────┬───────────────────────────────┘
                          │
                          ▼
┌─────────────────────────────────────────────────────────┐
│ EXCEL SPREADSHEET                                       │
│                                                         │
│ Лист "Игра":            Лист "История":                 │
│ ├─ Баланс (B3)          ├─ №                             │
│ ├─ Ставка (D3)          ├─ Время                         │
│ ├─ Кнопки (F-M)         ├─ Тип ставки                    │
│ └─ Результаты (B9)      ├─ Число                         │
│                         ├─ Цвет                         │
│                         ├─ Ставка                       │
│                         └─ Результат                    │
└─────────────────────────────────────────────────────────┘
```

---

## 🎓 Как расширить проект

### Добавить новый тип ставки

```vba
' 1. В RouletteGame.bas, функция CalculateWinnings():
Case "NEWSTYPE"
    If [условие выигрыша] Then
        payout = [коэффициент]
    Else
        payout = 0
    End If

' 2. В GameInterface.bas, новый обработчик:
Sub BtnNewType_Click()
    Dim betAmount As Currency
    betAmount = GetBetAmount()
    If betAmount > 0 Then
        PlaceBet betAmount, "NEWSTYPE", ""
    End If
End Sub

' 3. В SpreadsheetLayout.bas, в SetupGameSheet():
Call CreateButton(ws, "...", "...", "НОВАЯ СТАВКА", "BtnNewType", RGB(...), 12)
```

### Добавить новую стратегию

```vba
' 1. В RouletteGame.bas, новая функция:
Sub LoadStrategyNew()
    gameState.CurrentStrategy = "NewStrategy"
    ' Логика стратегии
End Sub

' 2. В GameInterface.bas, обработчик:
Sub BtnStrategyNew_Click()
    LoadStrategyNew
    MsgBox "Стратегия загружена: Новая", vbInformation
End Sub
```

### Добавить звуки

```vba
' В RouletteGame.bas, в функцию DisplaySpinResult():
If winnings > 0 Then
    PlaySound "C:\Windows\Media\tada.wav"  ' Выигрыш
Else
    PlaySound "C:\Windows\Media\chord.wav" ' Проигрыш
End If
```

---

## 📞 Контакты и ссылки

- **GitHub Репозиторий**: https://github.com/asers2022/Roulette-Game
- **Автор**: asers2022
- **Email**: asers2022@email.com (если указан)
- **Issues**: https://github.com/asers2022/Roulette-Game/issues

---

## ✅ Чек-лист файлов

- [x] README.md - основная документация
- [x] QUICKSTART.md - быстрый старт
- [x] INSTALLATION.md - подробная установка
- [x] PROJECT_STRUCTURE.md - этот файл
- [x] src/RouletteGame.bas - основная логика
- [x] src/GameInterface.bas - обработчики кнопок
- [x] src/SpreadsheetLayout.bas - интерфейс
- [ ] docs/ - дополнительные документы (можно добавить)
- [ ] examples/ - примеры (можно добавить)
- [ ] LICENSE - лицензия (обычно MIT)

---

**Версия**: 1.0  
**Дата**: 2026-07-19  
**Автор**: asers2022
