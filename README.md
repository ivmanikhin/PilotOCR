# PilotOCR инструкция по настройке
Файлы настроек connection_settings.txt и acceptable_doc_types.txt положить в "C:\Users\\{username}\AppData\Local\ASCON\Pilot-ICE Enterprise\PilotOCR\".
Acceptable doc types - типы документов, подлежащие распознаванию. connection_settings - настройки подключесния к базе SQL.
Для создания базы данных использовать Database.sql. Используется база данных MariaDB или MySQL. В папке с проектом необходимо создать папку Tessdata, положить туда файл
https://github.com/tesseract-ocr/tessdata/blob/main/rus.traineddata и сделать его "embedded resource"