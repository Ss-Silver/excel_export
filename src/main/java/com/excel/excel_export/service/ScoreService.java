package com.excel.excel_export.service;

import com.excel.excel_export.dto.ScoreDto;
import com.excel.excel_export.entity.Score;
import com.excel.excel_export.repository.ScoreRepository;
import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Service;

import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;

@Service
@RequiredArgsConstructor
public class ScoreService {
    private final ScoreRepository scoreRepository;

    public void setup() {

        Score score1 = Score.builder()
                .seq("1반")
                .name("김학생")
                .korean("90")
                .math("80")
                .english("70")
                .history("60")
                .birthDay(LocalDateTime.now())
                .money(10000)
                .build();
        Score score2 = Score.builder()
                .seq("2반")
                .name("나학생")
                .korean("90")
                .math("80")
                .english("70")
                .history("60")
                .birthDay(LocalDateTime.now())
                .money(1000)
                .build();
        Score score3 = Score.builder()
                .seq("3반")
                .name("다학생")
                .korean("90")
                .math("80")
                .english("70")
                .history("60")
                .birthDay(LocalDateTime.now())
                .money(100)
                .build();
        Score score4 = Score.builder()
                .seq("4반")
                .name("라학생")
                .korean("90")
                .math("80")
                .english("70")
                .history("60")
                .birthDay(LocalDateTime.now())
                .money(10)
                .build();

        scoreRepository.save(score1);
        scoreRepository.save(score2);
        scoreRepository.save(score3);
        scoreRepository.save(score4);
    }

    public List<ScoreDto> getScoreInfo(){
        List<Score> scores = scoreRepository.findAll();
        List<ScoreDto> ScoreDto = new ArrayList<>();

        for (Score score1 : scores) {
            ScoreDto scoreDto = new ScoreDto(score1);
            ScoreDto.add(scoreDto);
        }
        return ScoreDto;
    }

}
