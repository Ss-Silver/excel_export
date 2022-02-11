package com.excel.excel_export.dto;

import com.excel.excel_export.entity.Score;
import com.excel.excel_export.format.*;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import java.awt.*;
import java.math.BigDecimal;
import java.time.LocalDateTime;

@Setter
@Getter
@NoArgsConstructor
public class ScoreDto {
    @ExcelHeader(headerName = "반", colIndex = 0, rowIndex = 0, rowSpan = 1, headerStyle = @HeaderStyle(background = @Background(color = "#ECEFF3")))
    @ExcelBody(rowIndex = 0, colIndex = 0, rowGroup = true, bodyStyle = @BodyStyle(horizontalAlignment = HorizontalAlignment.CENTER))
    private String seq;
    @ExcelHeader(headerName = "이름", colIndex = 1, rowIndex = 0, rowSpan = 1, headerStyle = @HeaderStyle(background = @Background(color = "#ECEFF3")))
    @ExcelBody(rowIndex = 0, colIndex = 1, rowSpan = 1, bodyStyle = @BodyStyle(background = @Background(color = "#1E90FF")))
    private String name;

    @ExcelHeader(headerName = "국어/수학", colIndex = 2, colSpan = 1, rowIndex = 0, headerStyle = @HeaderStyle(background = @Background(color = "#ECEFF3")))
    private String koreanMathHeader;

    @ExcelBody(rowIndex = 0, colIndex = 2, bodyStyle = @BodyStyle(background = @Background(color = "#008080")))
    private String korean;
    @ExcelBody(rowIndex = 0, colIndex = 3)
    private String math;

    @ExcelHeader(headerName = "영어", colIndex = 2, rowIndex = 1, headerStyle = @HeaderStyle(background = @Background(color = "#ECEFF3")))
    @ExcelBody(rowIndex = 1, colIndex = 2)
    private String english;
    @ExcelHeader(headerName = "역사", colIndex = 3, rowIndex = 1, headerStyle = @HeaderStyle(background = @Background(color = "#ECEFF3")))
    @ExcelBody(rowIndex = 1, colIndex = 3)
    private String history;

    @ExcelHeader(headerName = "생일", colIndex = 4, rowIndex = 0, headerStyle = @HeaderStyle(background = @Background(color = "#ECEFF3")))
    @ExcelBody(rowIndex = 0, colIndex = 4, bodyStyle = @BodyStyle(dateFormat = "YYYY-MM-DD"), width = 12)
    private LocalDateTime birthDay;

    @ExcelHeader(headerName = "금액", colIndex = 4, rowIndex = 1, headerStyle = @HeaderStyle(background = @Background(color = "#ECEFF3")))
    @ExcelBody(rowIndex = 1, colIndex = 4, bodyStyle = @BodyStyle(numberFormat = "#,##0원"))
    private int money;

    public ScoreDto(Score score) {
        this.seq = score.getSeq();
        this.name = score.getName();
        this.korean = score.getKorean();
        this.math = score.getMath();
        this.english = score.getEnglish();
        this.history = score.getHistory();
        this.birthDay = score.getBirthDay();
        this.money = score.getMoney();
    }
}
