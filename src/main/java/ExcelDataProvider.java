/**
 * Data provider for ExcelReaderImpl
 * Created by AZiatdinov on 09.02.2017.
 */
// TODO: ужасное название для ячейки. Может просто Cell?
class ExcelDataProvider {
    private Object value;
    private int regionId;

    ExcelDataProvider(Object value, int regionId){
        this.value = value;
        this.regionId = regionId;
    }

    Object getValue(){
        return value;
    }

    int getRegionId(){
        return regionId;
    }
}
