package com.excel.excel_export.entity;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.Id;
import java.math.BigDecimal;
import java.time.LocalDateTime;

@Entity
@Getter
@Builder
@AllArgsConstructor
@NoArgsConstructor
public class Score {

    @Id @GeneratedValue
    @Column(name = "score_id")
    private Long id;

    private String seq;
    private String name;
    private String korean;
    private String math;
    private String english;
    private String history;
    private LocalDateTime birthDay;
    private int money;

}
