# Учёт рентгенологических исследований  

## Стек технологий  

![Static Badge](https://img.shields.io/badge/python-white?style=for-the-badge&logo=python&logoColor=%233776AB&labelColor=%23ebe3ce&color=%233776AB)  


## Установка проекта локально (Linux)  
+ Склонировать репозиторий и перейти в него в командной строке:  
```
git clone https://github.com/lobster2nd/report_counter.git  
cd report_counter 
```  
+ Cоздать и активировать виртуальное окружение:   
```
python -m venv env
```  
```
source env/bin/activate
```  
+ Перейти в директорию и установить зависимости из файла requirements.txt:  
```
pip install -r requirements.txt
```

+ Запуск приложения:  

```
flet run main.py  
```

+ Сборка приложения (по умолчанию --name принимает заничение указанного скрипта):  
  
```
pip install pyinstaller
flet pack main.py --name bundle_name
```

Приложение будет доступо в директории:  
+ Windows:
```
dist\your_program.exe
```
+ Linux:
```
dist/your_program
```
+ MacOS:
```
open dist/your_program.app
```
