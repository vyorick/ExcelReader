# ExcelReader
Обработчик Excel файлов

Содержит интерфейс со следующими методами:

    // Определяет файл для работы, принимая путь в качестве параметра
    void setExcelFile(String Path);

    // Переключает на страницу с указанным именем
    void switchToSheet(String SheetName);

    // Переключает на страницу с указанным номером
    void switchToSheet(int number);

    // Выводит таблицу данных
    Object[][] getData();

    // Выводит содержимое строки по её номеру
    Object[] getRowData(int rowNo);

    // Предоставляет содержимое ячейки
    Object getCellData(int rowNum, int colNum);
