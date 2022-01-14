package com.mmg;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

/**
 * @Auther: fan
 * @Date: 2021/12/31
 * @Description:
 */
public class Main {

    public static void main(String[] args) throws IOException {
        Main main = new Main();
//        System.out.println(main.createPDF(new ArrayList<>()));
        XWPFDocument document = new XWPFDocument(new FileInputStream("C:\\Users\\admin\\Desktop\\test.docx"));
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        for (XWPFParagraph paragraph : paragraphs) {
            for (XWPFRun run : paragraph.getRuns()) {
                System.out.println(run.getText(0) + "\n");
            }
        }

    }

    public String createPDF(List<ExecutionPdfVO> executionPdfVOS) throws IOException {
        //输入模板
        XWPFDocument document = new XWPFDocument(new FileInputStream("C:\\Users\\admin\\Desktop\\test.docx"));
        //获取表格
        List<XWPFTable> tables = document.getTables();
        //操作第一个表格
        XWPFTable table1 = tables.get(0);
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        lureInfos(executionPdfVOS).forEach(item -> {
            XWPFTableRow row = table1.createRow();
            List<XWPFTableCell> cells = row.getTableCells();
            //第几列
            try {
                insertCell(cells.get(2), sdf.format(item.getStartTime()) + " 至 " + sdf.format(item.getEndTime()));
                insertCell(cells.get(3), item.getExecutionResult());
            } catch (IOException e) {
                e.printStackTrace();
            }
        });
        //合并单元格
        List<Integer> cropList = new ArrayList<>();
        List<Integer> lureList = new ArrayList<>();
        for (ExecutionPdfVO executionPdfVO : executionPdfVOS) {
            int a = 0;
            for (ExecutionPdfVO.DataInfo dataInfo : executionPdfVO.getInfo()) {
                lureList.add(dataInfo.getLureInfoList().size());
                a = a + dataInfo.getLureInfoList().size();
            }
            cropList.add(a);
        }
        int lureStart = 1;
        for (Integer lureNum : lureList) {
            int lureEnd = lureStart + lureNum;
            mergeCellVertically(table1, 1, lureStart, lureEnd - 1);
            lureStart = lureEnd;
        }
        int cropStart = 1;
        for (Integer cropNum : cropList) {
            int cropEnd = cropStart + cropNum;
            mergeCellVertically(table1, 0, cropStart, cropEnd - 1);
            cropStart = cropEnd;
        }

        //操作第二个表格
        XWPFTable table2 = tables.get(1);

        //输出
        document.write(new FileOutputStream("C:\\Users\\admin\\Desktop\\new.docx"));
        return null;
    }

    //填充数据
    private void insertCell(XWPFTableCell cell, String text) throws IOException {
        //内容和样式
        XWPFRun run = cell.getParagraphs().get(0).createRun();
        run.setText(text);
        run.setFontSize(10);
        //居中
        CTTc ctTc = cell.getCTTc();
        CTTcPr ctPr = ctTc.addNewTcPr();
        ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);
        ctTc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.CENTER);
    }

    //获取所有最下层列表
    private List<ExecutionPdfVO.LureInfo> lureInfos(List<ExecutionPdfVO> executionPdfVOS) {
        List<ExecutionPdfVO.LureInfo> list = new ArrayList<>();
        executionPdfVOS.forEach(item -> item.getInfo().forEach(dataInfo -> list.addAll(dataInfo.getLureInfoList())));
        return list;
    }

    //合并单元格
    public static void mergeCellVertically(XWPFTable table, int col, int fromRow, int toRow) {
        for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
            CTVMerge merge = CTVMerge.Factory.newInstance();
            if (rowIndex == fromRow) {
                merge.setVal(STMerge.RESTART);
            } else {
                merge.setVal(STMerge.CONTINUE);
            }
            XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
            CTTcPr tcPr = cell.getCTTc().getTcPr();
            if (tcPr != null) {
                tcPr.setVMerge(merge);
            } else {
                tcPr = CTTcPr.Factory.newInstance();
                tcPr.setVMerge(merge);
                cell.getCTTc().setTcPr(tcPr);
            }
        }
    }
}
