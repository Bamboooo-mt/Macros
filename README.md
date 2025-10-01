## README for VBA Macros

This repository contains a set of VBA macros to automate common document editing tasks in Microsoft Word. The macros simplify operations such as adding footnotes, formatting images, converting text to hyperlinks, and more.

### Included Macros

| Macro Module Name                   | Description                                                                                  |
|-------------------------------------|----------------------------------------------------------------------------------------------|
| **addFootnotesToHackTools.bas**     | Automatically adds footnotes to specified tool names or keywords.                            |
| **applyImageStyleAndCaption.bas**   | Applies a predefined image style and inserts a caption for each image in the document.       |
| **CVEtoHyperlinks.bas**             | Finds CVE identifiers (e.g., `CVE-2021-34527`) and converts them to clickable hyperlinks.    |
| **deleteReadyComments.bas**         | Deletes all resolved (finished / greyed out) comments in the document.                       |
| **exportHotKeys.bas**               | Exports the list of custom keyboard shortcuts (hotkeys) defined in Word to an external file. |
| **findRepeated.bas**                | Searches for repeated words or phrases and highlights them for review.                       |
| **keepWithNext.bas**                | Sets the ‘Keep With Next’ paragraph formatting for the selected paragraph.                   |
| **LinksToFootnotes.bas**            | Converts all hyperlinks to numbered footnotes in a selected range of text.                   |
| **maskingPass.bas**                 | Mask sensitive text, such as passwords, in a document according to a specified pattern.      |
| **replaceImagesFromFolder.bas**     | Replaces images in the document with files from the specified folder by file caption name.   |

### Prerequisites

- Microsoft Word (Office 2010 or later).
- Basic familiarity with the VBA editor (Alt + F11).

### Installation

1. Open your Word document.
2. Press **Alt + F11** to open the VBA editor.
3. In the Project pane, right-click **Modules** and choose **Import File...**.
4. Select the desired `.bas` files from this repository.
5. Close the VBA editor.

### Usage

1. In Word, go to the **View** tab, click **Macros**, then **View Macros**.
2. Select the desired macro from the list (e.g., `addFootnotesToHackTools`).
3. Click **Run**.
4. Follow any on-screen prompts (if the macro requires additional input).

To assign a keyboard shortcut:
1. Go to **File** > **Options** > **Customize Ribbon**.
2. Click **Customize...** next to **Keyboard shortcuts**.
3. In the **Categories** list, select **Macros**.
4. In the **Macros** list, select your macro.
5. Press the desired shortcut key, then click **Assign**.

## Руководство для макросов VBA

В данном репозитории собраны макросы VBA для автоматизации часто выполняемых задач в Microsoft Word. Макросы упрощают операции, такие как добавление сносок, форматирование изображений, преобразование текста в гиперссылки и многое другое.

### Список макросов

| Имя модуля                          | Описание                                                                                      |
|-------------------------------------|-----------------------------------------------------------------------------------------------|
| **addFootnotesToHackTools.bas**     | Автоматически добавляет сноски к заданным названиям инструментов или ключевым словам.         |
| **applyImageStyleAndCaption.bas**   | Применяет предустановленный стиль изображения и вставляет подпись к каждому изображению.      |
| **CVEtoHyperlinks.bas**             | Находит идентификаторы CVE (например, `CVE-2021-34527`) и делает их кликабельными ссылками.   |
| **deleteReadyComments.bas**         | Удаляет все решенные (готовые / серые) комментарии в документе.                               |
| **exportHotKeys.bas**               | Экспортирует список пользовательских сочетаний клавиш, определённых в Word, во внешний файл.  |
| **findRepeated.bas**                | Ищет повторяющиеся слова или фразы и выделяет их для проверки.                                |
| **keepWithNext.bas**                | Устанавливает форматирование абзаца «Не отрывать от следюущего» для выбранного абзаца.        |
| **LinksToFootnotes.bas**            | Преобразует все гиперссылки в нумерованные сноски в выбранном диапазоне текста.               |
| **maskingPass.bas**                 | Маскирует по заданному паттерну конфиденциальный текст, например пароли, в документе.         |
| **replaceImagesFromFolder.bas**     | Заменяет изображения в документе файлами из указанной папки по имени подписи файла.           |

### Требования

- Microsoft Word (Office 2010 или новее).
- Базовые навыки работы с редактором VBA (Alt + F11).

### Установка

1. Откройте документ Word.
2. Нажмите **Alt + F11** для открытия редактора VBA.
3. В панели проекта щелкните правой кнопкой по **Modules** и выберите **Import File...**.
4. Выберите нужные файлы `.bas` из этого репозитория.
5. Закройте редактор VBA.

### Использование

1. В Word перейдите на вкладку **View**, нажмите **Macros**, затем **View Macros**.
2. Выберите нужный макрос из списка (например, `addFootnotesToHackTools`).
3. Нажмите **Run**.
4. Следуйте инструкциям на экране (если макрос требует дополнительного ввода).

Чтобы назначить сочетание клавиш:
1. Перейдите в **File** > **Options** > **Customize Ribbon**.
2. Нажмите **Customize...** рядом с **Keyboard shortcuts**.
3. В списке **Categories** выберите **Macros**.
4. В списке **Macros** выберите ваш макрос.
5. Нажмите желаемую клавишу, затем **Assign**.
# Macros
