package cn.isuyu.poi.operate.word.dtos;

import lombok.Data;

import java.io.Serializable;
import java.util.List;

/**
 * @Author NieZhiLiang
 * @Email nzlsgg@163.com
 * @GitHub https://github.com/niezhiliang
 * @Date 2020/5/7 下午6:59
 */
@Data
public class OperateDTO implements Serializable {

    List<DynamicTableDTO> dynamicRows;
}
