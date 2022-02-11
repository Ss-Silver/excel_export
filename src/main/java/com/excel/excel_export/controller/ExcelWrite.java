package com.excel.excel_export.controller;

import com.excel.excel_export.format.ExcelBody;
import com.excel.excel_export.format.ExcelHeader;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.awt.*;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.net.URLEncoder;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;
import java.util.List;
import java.util.stream.Collectors;

/**
 *
 *
 * */

@RestController
public class ExcelWrite {

    /**
     * 매개변수는 파일로 출력될 엑셀파일명, 출력 클래스, 출력 데이터 를 전달하여준다.
     * */
    public static <T> ResponseEntity<ByteArrayResource> export(String fileName, Class<T> excelClass, List<T> data) throws IllegalAccessException, IOException {

        /**
         * 엑셀 파일 및 시트를 생성하여 준다.
         * Sheet 명은 현재는 fileName과 같다. 만약 sheetName을 따로 설정하고 싶은경우는
         * 새로운 매개변수를 생성 및 받아와 삽입 혹은 지정된 String 값을 아래
         * createSheet( {SheetName} ); 에 넣어준다.
         * */
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet(fileName);

        /**
         * Map을 사용하여 @ExcelHeader의 값을 rowIndex를 key값으로 하여 Map을 구성한다.
         * */
        Map<Integer, List<ExcelHeader>> headerMap = Arrays.stream(excelClass.getDeclaredFields())
                .filter(field -> field.isAnnotationPresent(ExcelHeader.class))
                .map(field -> field.getDeclaredAnnotation(ExcelHeader.class))
                .sorted(Comparator.comparing(ExcelHeader::colIndex))
                .collect(Collectors.groupingBy(ExcelHeader::rowIndex));

        /**
        * 위 헤더를 담은 Map에서 설정값들을 빼와 자동으로 set을 시켜준다.
         * (현재버전 배경색을 설정할 수 있다._22.02.11 성상은)
       *  현재는 Border 고정으로 4방면 실선으로 입력해있다.
         * 추후에는 BorderStyle을 따로 set하는 것을 고민하고 적용을 해봐야한다.
        * */
        int index = 0;
        for(Integer key : headerMap.keySet()) {
            XSSFRow row = sheet.createRow(index++);

            for (ExcelHeader excelHeader : headerMap.get(key)) {
                XSSFCell cell = row.createCell(excelHeader.colIndex());
                XSSFCellStyle cellStyle = workbook.createCellStyle();
                cell.setCellValue(excelHeader.headerName());
                if(excelHeader.headerName().contains("\n")) {
                    cellStyle.setWrapText(true);
                }
                cellStyle.setAlignment(excelHeader.headerStyle().horizontalAlignment());
                cellStyle.setVerticalAlignment(excelHeader.headerStyle().verticalAlignment());
                if(excelHeader.headerStyle().background().color() != null) {
                    cellStyle.setFillForegroundColor(new XSSFColor(Color.decode(excelHeader.headerStyle().background().color()), new DefaultIndexedColorMap()));
                    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                }
                XSSFFont font = workbook.createFont(); //기본 font값은 'Calibri' 이다.
                font.setFontHeightInPoints((short) excelHeader.headerStyle().fontSize());
                cellStyle.setFont(font);
                cellStyle.setBorderBottom(BorderStyle.THIN);
                cellStyle.setBorderLeft(BorderStyle.THIN);
                cellStyle.setBorderRight(BorderStyle.THIN);
                cellStyle.setBorderTop(BorderStyle.THIN);
                cell.setCellStyle(cellStyle);
                if(excelHeader.colSpan() > 0 || excelHeader.rowSpan() > 0) {
                    CellRangeAddress cellAddresses =
                            new CellRangeAddress(cell.getAddress().getRow(), cell.getAddress().getRow() + excelHeader.rowSpan(),
                                    cell.getAddress().getColumn(), cell.getAddress().getColumn() + excelHeader.colSpan());
                    sheet.addMergedRegion(cellAddresses);
                }
            }
        }

        /**
         * Map을 사용하여 @ExcelBody의 값을 rowIndex를 key값으로 하여 Map을 구성한다.
         * */
        Map<Integer, List<Field>> fieldMap = Arrays.stream(excelClass.getDeclaredFields())
                .filter(field -> field.isAnnotationPresent(ExcelBody.class))
                .map(field -> {
                    field.setAccessible(true);
                    return field;
                })
                .sorted(Comparator.comparing(field -> field.getDeclaredAnnotation(ExcelBody.class).colIndex()))
                .collect(Collectors.groupingBy(field -> field.getDeclaredAnnotation(ExcelBody.class).rowIndex()));

        /**
         * 위 Body를 담은 Map에서 설정값들을 빼와 자동으로 set을 시켜준다.
         * (현재버전 배경색을 설정할 수 있다._22.02.11 성상은)
         *  Body의 경우 수직정렬은 Center로 설정되어있지만, 수평 정렬은 General 이다.
         *  현재는 Border 고정으로 4방면 실선으로 입력해있다.
         * 추후에는 BorderStyle을 따로 set하는 것을 고민하고 적용을 해봐야한다.
         * */
        for (T t : data) {
            for (Integer key : fieldMap.keySet()) {
                XSSFRow row = sheet.createRow(index++);
                for (Field field : fieldMap.get(key)) {
                    ExcelBody excelBody = field.getDeclaredAnnotation(ExcelBody.class);
                    Object o = field.get(t);
                    XSSFCell cell = row.createCell(excelBody.colIndex());
                    XSSFCellStyle cellStyle = workbook.createCellStyle();
                    XSSFDataFormat dataFormat = workbook.createDataFormat();

                    cellStyle.setAlignment(excelBody.bodyStyle().horizontalAlignment());
                    cellStyle.setVerticalAlignment(excelBody.bodyStyle().verticalAlignment());
                    if (excelBody.bodyStyle().background().color() != null) {
                        cellStyle.setFillForegroundColor(new XSSFColor(Color.decode(excelBody.bodyStyle().background().color()), new DefaultIndexedColorMap()));
                        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    }
                    cellStyle.setBorderBottom(BorderStyle.THIN);
                    cellStyle.setBorderLeft(BorderStyle.THIN);
                    cellStyle.setBorderRight(BorderStyle.THIN);
                    cellStyle.setBorderTop(BorderStyle.THIN);

                    if (o instanceof Number) {
                        if (excelBody.bodyStyle().numberFormat() != null) {
                            cellStyle.setDataFormat(dataFormat.getFormat(excelBody.bodyStyle().numberFormat()));
                        }
                        cell.setCellValue(((Number) o).doubleValue());
                    } else if (o instanceof String) {
                        cell.setCellValue((String) o);
                    } else if (o instanceof Date) {
                        cellStyle.setDataFormat(dataFormat.getFormat(excelBody.bodyStyle().dateFormat()));
                        cell.setCellValue((Date) o);
                    } else if (o instanceof LocalDateTime) {
                        cellStyle.setDataFormat(dataFormat.getFormat(excelBody.bodyStyle().dateFormat()));
                        cell.setCellValue((LocalDateTime) o);
                    } else if (o instanceof LocalDate) {
                        cellStyle.setDataFormat(dataFormat.getFormat(excelBody.bodyStyle().dateFormat()));
                        cell.setCellValue((LocalDate) o);
                    }
                    cell.setCellStyle(cellStyle);
                    if (excelBody.colSpan() > 0 || excelBody.rowSpan() > 0) {
                        CellRangeAddress cellAddresses = new CellRangeAddress(cell.getAddress().getRow(), cell.getAddress().getRow() + excelBody.rowSpan(), cell.getAddress().getColumn(), cell.getAddress().getColumn() + excelBody.colSpan());
                        sheet.addMergedRegion(cellAddresses);
                    }
                    if ((excelBody.width() > 0 && excelBody.width() != 8) && sheet.getColumnWidth(excelBody.colIndex()) == 2048) {
                        sheet.setColumnWidth(excelBody.colIndex(), excelBody.width() * 256);
                    }
                }
            }
        }

        /**
         * 위 기능은 rowGroup 속성으로 해당 설정이 존재하는 필드를 가져온다.
         * -> 이는 특정 Column의 이전 행과 현재 행 값을 비교하여 동적으로 Column을 병합하는 기능에 사용됨.
         * */
        List<Field> groupField = Arrays.stream(excelClass.getDeclaredFields())
                .filter(field -> field.isAnnotationPresent(ExcelBody.class) && field.getDeclaredAnnotation(ExcelBody.class).rowGroup())
                .map(field -> {
                    field.setAccessible(true);
                    return field;
                })
                .sorted(Comparator.comparing(field -> field.getDeclaredAnnotation(ExcelBody.class).colIndex()))
                .collect(Collectors.toList());

        /**
         * 전체 data를 반복하여 현재행과 다음 행의 칼럼을 비교하며 값이 다른 인덱스를 탐색한다.
         * */
        Map<Field, List<Integer>> groupMap = new HashMap<>();
        for (Field field : groupField){
            groupMap.put(field, new ArrayList<>());
            for(int i=0; i< data.size(); i++){
                Object o1 = field.get(data.get(i));

                for(int j = i+1;j < data.size(); j++){
                    Object o2 = field.get(data.get(j));
                    if(!o1.equals(o2)){
                        groupMap.get(field).add((j)* headerMap.size()+headerMap.keySet().size()-1);
                        i = j-1;
                        break;
                    }
                }
            }
            groupMap.get(field).add(sheet.getLastRowNum());
        }


        /**
         * 탐색한 인덱스를 이용하여 데이터행 부터 rowGorup이 설정된 필드의 동일한 값을 갖는 칼럼을 병합한다
         * */
        for(Field field: groupMap.keySet()){
            int dataRowIndex = headerMap.keySet().size();
            for(int i=0; i<groupMap.get(field).size(); i++){
                XSSFRow row = sheet.getRow(dataRowIndex);
                XSSFCell cell = row.getCell(field.getDeclaredAnnotation(ExcelBody.class).colIndex());
                if(!(dataRowIndex == groupMap.get(field).get(i))){
                    CellRangeAddress cellAddresses = new CellRangeAddress(dataRowIndex, groupMap.get(field).get(i), cell.getColumnIndex(), cell.getColumnIndex());
                    sheet.addMergedRegion(cellAddresses);
                }
                dataRowIndex = groupMap.get(field).get(i)+1;
            }
        }


        /**
         * 행 병합, 열 병합한 칼럼들의 테두리를 재 설정해줌
         * */
        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
        for(CellRangeAddress rangeAddress : mergedRegions) {
            RegionUtil.setBorderBottom(BorderStyle.THIN, rangeAddress, sheet);
            RegionUtil.setBorderLeft(BorderStyle.THIN, rangeAddress, sheet);
            RegionUtil.setBorderRight(BorderStyle.THIN, rangeAddress, sheet);
            RegionUtil.setBorderTop(BorderStyle.THIN, rangeAddress, sheet);
        }


        /**
         * 일반적으로 OutputStream 사용하여 byte[]받아 write를 사용한다.
         * 하지만, 바이트 기반 스트림의 기능을 보완(성능향상 및 기능추가)하기 위하여
         * ByteArrayOutputStream 을 사용한다. OutputStream 을 상속받으며 입출력기능은 위임한다.
         *
         * 참고자료 :
         * https://joont92.github.io/java/%EC%9E%85%EC%B6%9C%EB%A0%A5/
         * */
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        workbook.write(byteArrayOutputStream);

        return ResponseEntity
                .ok()
                .header("Content-Transfer-Encoding", "binary")
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + URLEncoder.encode(fileName, "UTF-8")+".xlsx")
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .contentLength(byteArrayOutputStream.size())
                .body(new ByteArrayResource(byteArrayOutputStream.toByteArray()));
    }


}
