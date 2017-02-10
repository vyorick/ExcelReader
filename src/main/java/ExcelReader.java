/**
 * Interface for ExcelReaderImpl
 * Created by AZiatdinov on 08.02.2017.
 */
interface ExcelReader {

    // Определяет файл для работы, принимая путь в качестве параметра
    void setExcelFile(String Path);

    // Переключает на страницу с указанным именем
    void switchToSheet(String SheetName);

    // Переключает на страницу с указанным номером
    void switchToSheet(int number);

    // Выводит таблицу данных
    // TODO: правильно ли я понимаю, что длина вложенных масивов различна? Если "да", то лучше указать этот момент
    // TODO: getData => getCurrentSheetData
    // TODO: может тогда сделаем обертку: getSheetData(String sheetName)?
    Object[][] getData();

    // TODO: не выводит, а возвращает
    // Выводит содержимое строки по её номеру
    Object[] getRowData(int rowNo);

    // Предоставляет содержимое ячейки
    Object getCellData(int rowNum, int colNum);

}
