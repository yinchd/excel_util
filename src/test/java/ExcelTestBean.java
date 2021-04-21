import com.yinchd.excel.ExcelField;
import com.yinchd.excel.ExcelSheet;
import lombok.Data;

@Data
@ExcelSheet()
public class ExcelTestBean {

    @ExcelField(name = "BMSAH")
    String bmsah;

    @ExcelField(name = "TYSAH")
    String tysah;

    @ExcelField(name = "XYRBH")
    String xyrbh;

    @ExcelField(name = "XM")
    String xm;

    @ExcelField(name = "ZJLX_MC")
    String zjlxMc;

    @ExcelField(name = "ZJHM")
    String zjhm;
}
