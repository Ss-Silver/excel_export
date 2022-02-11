package com.excel.excel_export.format;

import org.apache.poi.ss.usermodel.Color;

import java.awt.*;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface Background {
    String color() default "#f5f5dc";
}
