package com.excel.excel_export.format;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelHeader {
    String headerName();
    int rowIndex() default 0;
    int colIndex();
    int colSpan() default 0;
    int rowSpan() default 0;
    HeaderStyle headerStyle() default @HeaderStyle;

}
