import com.alibaba.fastjson.JSON;
import com.yinchd.excel.ExcelReader;
import com.yinchd.excel.ExcelWriter;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.time.Duration;
import java.time.LocalTime;
import java.util.List;

public class Main {

    public static void main(String[] args) throws FileNotFoundException {
        String filePath = "D:\\Download\\test.xlsx";
        FileInputStream fis = new FileInputStream(filePath);
        LocalTime t1 = LocalTime.now();
        List<ExcelTestBean> all = ExcelReader.getListByFilePath(filePath, ExcelTestBean.class);
        System.out.println("表格的详细数据为：" + JSON.toJSONString(all));
        System.out.println("all的大小为：" + all.size());
        System.out.println("读取整个表格耗时：" + Duration.between(t1, LocalTime.now()).getSeconds() + "s");
        ExcelWriter.writeToDesktop("导出测试", all);

    }
}
