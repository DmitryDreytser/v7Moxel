# v7Moxel

Внешняя компонента и класс для конвертации печатных таблиц 1C 7.7

Сейчас умеет читать MOXEL версии 6 и 7, читать и отрисовывать OLE объекты (за исключением диаграмм) в *.png с прозрачностью и сохранять в HTML, XLSX и PDF

В PDF пока что сохраняет только в ландшафтной ориентации. Объект с настройками страницы еще не реализован.

Поддерживает BMP_1C (есть нюанс, OLE объект хранит в себе толоко путь к исходному изображению, поэтому работает только если оно на месте).

Поддерживает встроенные объекты Office и пр. все что может само отрисовать картинку.

Реализован интерфейс внешней компоненты с поддержкой ILanguageExtender для использования внутри 1С

Не реализовано:
- чтение формата версии 8 и пр.
- отрисовка диаграмм - OLE объект при загрузке вываливает дилог без текста.
- запись в mxl

