import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.util.List;
import java.util.logging.Logger;

/**
 * This is class for work with Excel files
 * Created by AZiatdinov on 08.02.2017.
 */

class ExcelReaderImpl implements ExcelReader {
    // TODO: Общее замечание: если ты кладешь в результат экземпляры конкрентного класса, то нужно переделать c Object на
    // TODO: этот класс. Это обеспечит как контроль, так и удобство использования (не придется приводить к типу)
    private XSSFWorkbook excelWBook;
    private XSSFSheet excelWSheet;

    private List<CellRangeAddress> mergedRegions;

    private Logger log = Logger.getLogger(this.getClass().getName());


    // Определяет файл для работы
    // TODO: коммент неудачный.
    public void setExcelFile(String Path) {
        try {
            FileInputStream ExcelFile = new FileInputStream(Path);
            excelWBook = new XSSFWorkbook(ExcelFile);
            ExcelFile.close();
        } catch (Exception e) {
            log.info("Can't open file " + e.getMessage());
        }
    }


    // Переключается на вкладку с указанным именем
    // TODO: Все аналогичные тудушки из метода ниже
    public void switchToSheet(String SheetName) {
        try {
            int sheetIndex = excelWBook.getSheetIndex(SheetName);
            switchToSheet(sheetIndex);
        } catch (Exception e) {
            log.info("Can't open sheet with name " + SheetName + " " + e.getMessage());
        }
    }


    // Переключается на вкладку с указанным номером
    // TODO: не нужно здесь отлавливать Exception, не принимай ответственность за косячно указанный "снаружи" номер страницы на внутренний класс.
    // TODO: Если нет страницы с таким номером, об этом нужно уведоммить, а не "тихонько" вывести в лог и не произвести ожидаемого действия,
    // TODO: иначе получим
    // TODO: 1. трудноуловимую ошибку
    // TODO: 2. невозможность определить произошло ли фактическое переключение на вкладку из вызывающего класса
    public void switchToSheet(int number) {
        try {
            excelWSheet = excelWBook.getSheetAt(number);
            mergedRegions = excelWSheet.getMergedRegions();
        } catch (Exception e) {
            log.info("Can't open sheet with number " + number + " " + e.getMessage());
        }
    }


    // Передаёт содержимое страницы
    public Object[][] getData() {
        int usedRows = getRowsUsed();
        Object[][] data = new Object[usedRows][];
        for (int i = 0; i < usedRows; i++){
            data[i] = getRowData(i);
        }
        return data;
    }


    // Выдаёт содержимое строки
    public Object[] getRowData(int rowNo) {
        int usedColumns = getColumnsUsed(rowNo);
        Object[] rowData = new Object[usedColumns];
        for (int i = 0; i < usedColumns; i++){
            rowData[i] =  getCellData(rowNo, i);
        }
        return rowData;
    }


    // Выводит данные ячейки
    // TODO: оно у тебя выводит экземпляр вполне конкретного класса - давай переделаем возвращаемый тип на реальный
    public Object getCellData(int rowNum, int colNum) {
        int region = getMergedRegion(rowNum, colNum);
        XSSFCell cell;
        try {
            if (region >= 0){
                cell = getMergedRegionStringValue(rowNum, colNum);
                // TODO: вынести из if, ибо дублируется
                return getStringValueFromCell(region, cell);
            } else {
                cell = excelWSheet.getRow(rowNum).getCell(colNum);
                // TODO: вынести из if (удалить)
                return getStringValueFromCell(region, cell);
            } //TODO: после удаления переделать на тернарный оператор
        } catch (Exception e){
            log.info("Can't read cell data " + e.getMessage());
            return null;
        }
    }


    // Получает значение ячейки типа String
    private Object getStringValueFromCell(int region, XSSFCell cell) {
        cell.setCellType(CellType.STRING);
        return new ExcelDataProvider(cell.getStringCellValue(), region);
    }


    // Выводит количество использованных строк
    private int getRowsUsed() {
        if (excelWSheet == null) {
            return 0;
        }
        return excelWSheet.getLastRowNum();
    }


    // Определяет количество использованных колонок в строке
    private int getColumnsUsed(int rowNo) {
        if (excelWSheet == null) {
            return 0;
        }
        return excelWSheet.getRow(rowNo).getPhysicalNumberOfCells();
    }


    // Определение номера региона, если ячейка объединена, по её координатам
    private int getMergedRegion(int rowNum, int colNum) {
        for (int i=0; i < mergedRegions.size(); i++) {
            if (mergedRegions.get(i).isInRange(rowNum, colNum)) {
                return i;
            }
        }
        return -1;
    }


    // Получение значений объединённых ячеек
    private XSSFCell getMergedRegionStringValue(int row, int column){
        int mergedRegionNumber = getMergedRegion(row, column);
        CellRangeAddress region = excelWSheet.getMergedRegion(mergedRegionNumber);

        int firstRegionColumn = region.getFirstColumn();
        int firstRegionRow = region.getFirstRow();

        return excelWSheet.getRow(firstRegionRow).getCell(firstRegionColumn);
    }
}
