import org.junit.Assert;
import org.junit.Test;

import java.util.logging.Logger;

/**
 * Test for ExcelReader
 * Created by AZiatdinov on 08.02.2017.
 */
public class ExcelReaderTest {

    private Logger log = Logger.getLogger(this.getClass().getName());
    private ExcelReader excelReader = new ExcelReaderImpl();

    @Test
    public void test(){
        String path = ".\\src\\main\\resources\\Table.xlsx";

        excelReader.setExcelFile(path);
        excelReader.switchToSheet(123); // Check error
        excelReader.switchToSheet("Main");

        Object[][] dataTable = excelReader.getData();
        for (Object[] rowData: dataTable){
            for (Object cellData: rowData){
                ExcelDataProvider data = (ExcelDataProvider) cellData;
                Assert.assertNotNull("Can't read file " + path + " ", data);
                log.info("Table value is: " + String.valueOf(data.getValue()) + " regionId: " + data.getRegionId());
            }
        }

        Object[] rowData = excelReader.getRowData(1);
        for (Object cellData: rowData){
            ExcelDataProvider data = (ExcelDataProvider) cellData;
            log.info("Row value is: " + String.valueOf(data.getValue()) + " regionId: " + data.getRegionId());
        }

        Object cellData = excelReader.getCellData(1, 1);
        ExcelDataProvider data = (ExcelDataProvider) cellData;
        log.info("Cell value is: " + String.valueOf(data.getValue()) + " regionId: " + data.getRegionId());
    }
}
