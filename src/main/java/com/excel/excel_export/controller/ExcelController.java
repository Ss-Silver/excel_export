package com.excel.excel_export.controller;

import com.excel.excel_export.dto.ScoreDto;
import com.excel.excel_export.service.ScoreService;
import lombok.RequiredArgsConstructor;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import static com.excel.excel_export.controller.ExcelWrite.*;

@RestController
@RequiredArgsConstructor
public class ExcelController {
    private final ScoreService scoreService;

    @PostMapping("/score")
    public List<ScoreDto> input() {
        scoreService.setup();
        List<ScoreDto> scoreDtos = scoreService.getScoreInfo();
        return scoreDtos;
    }

    @GetMapping("/excel")
    public ResponseEntity<ByteArrayResource> excel() throws IOException, IllegalAccessException {
        List<ScoreDto> scoreDtos = scoreService.getScoreInfo();
        ResponseEntity<ByteArrayResource> resource = export("test", ScoreDto.class, scoreDtos);
        return resource;

    }



}
