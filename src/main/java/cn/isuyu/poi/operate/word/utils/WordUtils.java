package cn.isuyu.poi.operate.word.utils;

import cn.isuyu.poi.operate.word.dtos.DynamicTableDTO;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;

import java.util.List;

/**
 * @Author NieZhiLiang
 * @Email nzlsgg@163.com
 * @GitHub https://github.com/niezhiliang
 * @Date 2020/5/7 下午5:06
 */
public class WordUtils {

    /**
     * 表格动态添加内容
     * @param xwpfDocument
     * @param dynamicTableDTOS
     */
    public static void dynamicTableReplace(XWPFDocument xwpfDocument, List<DynamicTableDTO> dynamicTableDTOS){
        try {
            //获取word文档中所有的表格
            List<XWPFTable> tables = xwpfDocument.getTables();
            if(tables !=null && !tables.isEmpty()) {
                //遍历操作的表单元格
                for (DynamicTableDTO tableDTO : dynamicTableDTOS) {
                    //获取操作的table
                    XWPFTable table = tables.get(tableDTO.getTableIndex());
                    //获取需要在哪一行下面插入数据
                    XWPFTableRow row1 = table.getRow(tableDTO.getRowIndex());
                    //得到需要填充的内容
                    List<List<String>> words = tableDTO.getRows();
                    int i = 0;
                    for(List<String> rowWord : words){
                        //直接将上一行的样式作为新的行的样式
                        table.addRow(row1,tableDTO.getRowIndex()+i);
                        //得到要添加的行
                        XWPFTableRow row = table.getRow(tableDTO.getRowIndex() + i);
                        //这里设置一个累加的变量是为了区分列的索引下标
                        int m = 0;
                        for (String value : rowWord) {
                            //获取需要填充的单元格
                            XWPFTableCell cell = row.getCell(m);
                            //设置单元格内容
                            setTableCell(cell,value,tableDTO);
                            m++;
                        }
                        i++;
                    }
                }
            }
        }
        catch (Exception e)
        {
            System.out.println(e.toString());
        }
    }

    /**
     * 设置单元格的值
     * @param cell
     * @param word
     * @param dynamicTableDTO
     */
    private static void setTableCell(XWPFTableCell cell,String word,DynamicTableDTO dynamicTableDTO){
        //   CTTcBorders tblBorders = cell.getCTTc().getTcPr().getTcBorders();

        //    CTBorder hBorder=tblBorders.addNewInsideH();
        //    hBorder.setVal(STBorder.Enum.forString("none"));
        //     hBorder.setSz(new BigInteger("1"));
        //    hBorder.setColor("0000FF");



        //    CTTcBorders tblBorders;
        //     tblBorders.addNewLeft().setVal(STBorder.NIL);
        //   tblBorders.addNewRight().setVal(STBorder.NIL);

        //   cell.getCTTc().getTcPr().setTcBorders(CTTcBorders.);


        cell.getParagraphs().get(0).removeRun(0);
        XWPFRun cellR = cell.getParagraphs().get(0).createRun();
        //单元格内容
        cellR.setText(word);
        //字体颜色
        cellR.setColor(dynamicTableDTO.getColor());
        //字体大小
        cellR.setFontSize(dynamicTableDTO.getFontSize());
        cellR.setFontFamily("仿宋");

        CTTc cttc = cell.getCTTc();
        CTTcPr ctPr = cttc.addNewTcPr();
        //默认单元格上下居中
        ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);

//        String fontalign = tkw.getFontalign();
//        if(fontalign!=null){
//
//            switch(fontalign)
//            {
//                case "left" :
//                    cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.LEFT);
//                    break;
//                case "right" :
//                    cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.RIGHT);
//                    break;
//                case "top" :
//                    ctPr.addNewVAlign().setVal(STVerticalJc.TOP);
//                    break;
//                case "bottom" :
//                    ctPr.addNewVAlign().setVal(STVerticalJc.BOTTOM);
//                    break;
//                case "lt" :
//                    cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.LEFT);
//                    ctPr.addNewVAlign().setVal(STVerticalJc.TOP);
//                    break;
//                case "rt" :
//                    cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.RIGHT);
//                    ctPr.addNewVAlign().setVal(STVerticalJc.TOP);
//                    break;
//                case "lb" :
//                    cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.LEFT);
//                    ctPr.addNewVAlign().setVal(STVerticalJc.BOTTOM);
//                    break;
//                case "rb" :
//                    cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.RIGHT);
//                    ctPr.addNewVAlign().setVal(STVerticalJc.BOTTOM);
//                    break;
//                case "hcenter" :
//                    cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.CENTER);
//                    break;
//                case "vcenter" :
//                    ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);
//                    break;
//                case "center" :
//                    cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.CENTER);
//                    ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);
//                    break;
//            }
//        }
    }
}
