# certificates-generator
## created by Anna Ilina in May 2021
Генератор получает на вход:

- **текстовый файл .txt**, в котором на каждой строке содержится одно имя, для которого необходимо сгенерировать сертификат (пример такого файла лежит в корне: names.txt),
- **файл шаблона с сертификатом в формате .docx**, в котором содержится переменная ***{{ name }}***, которая будет автоматически заменена на нужное имя (пример шаблона template.docx лежит в папке docx-templates).  

Для каждого имени из списка в файле .txt автоматически создастся сертификат с используемым шаблоном в формате .DOCX (и .PDF, если Вы работаете под ОС Windows, поскольку библиотеке необходим установленный Microsoft Office Word).  

Созданные сертификаты (DOCX и PDF) будут лежать в папке **certificates**.  

**В общем виде команда запуска скрипта должна выглядеть следующим образом:**  
```python3 main.py название_файла_со_списком_имён(напр.names.txt) название_файла_шаблона_docx(напр.docx-templates/template.docx)```
______________________________________________________________________________________________
### Инструкция по запуску программы на Windows:
1. Установить Python3
2. Скачать всю папку, разархивировать  (или клонировать из git)
3. Открыть командную строку от имени Администратора и последовательно ввести команды:
```
cd путь_к_certificates-generator-main_начиная_с_диска_C
```
```
gen-venv-windows\Scripts\activate.bat
```
```
py main.py names.txt docx-templates/template.docx
```
### Инструкция по запуску программы на Linux:
1. Установить Python3
2. Скачать всю папку, разархивировать (или клонировать из git)
3. В терминале последовательно выполнить следующие команды:
```
cd путь_к_certificates-generator-main
```
```
source gen-venv/bin/activate
```
```
python3 main.py names.txt docx-templates/template.docx
```
_____________________________________________________________________________________________
### Как создать шаблон DOCX из PPTX
1. Откройте pptx-шаблон, выделите всё (Crtl+A), правой кнопкой мыши -> Сгруппировать.
2. На сгруппированном списке нажмите правой кнопкой мыши -> Сохранить как рисунок. Сохраните в формате PNG.
3. Создайте пустой DOCX-документ и вставьте в него полученный PNG-рисунок.
4. Добавьте текстовое поле с текстом {{ name }} в место, где должно быть имя. Этот текст ( {{ name }} ) отформатируйте так, как должно выглядеть имя на конечном сертификате.
5. Сохраните в формате DOCX и запустите программу, выбрав полученный файл в качестве шаблона (прописать название этого файла последним аргументом).
