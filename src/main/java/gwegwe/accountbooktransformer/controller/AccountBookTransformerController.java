package gwegwe.accountbooktransformer.controller;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

@RestController
public class AccountBookTransformerController {

    @PostMapping("/process")
    public ResponseEntity<byte[]> processExcel(
            @RequestParam MultipartFile file,
            @RequestParam String startDate,
            @RequestParam String endDate,
            @RequestParam String owner
    ) {
        try {
            // 업로드된 엑셀 파일을 읽기 위해 Workbook 객체 생성
            Workbook workbook = WorkbookFactory.create(file.getInputStream());

            // 엑셀 시트에서 데이터를 추출
            List<List<String>> data = new ArrayList<>();
            Sheet sheet = workbook.getSheetAt(1); // 첫 번째 시트를 읽음
            Iterator<Row> rowIterator = sheet.iterator();

            // 목차 날리기.
            rowIterator.next();

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.iterator();
                List<String> rowData = new ArrayList<>();

                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    if (cell.getCellType() == CellType.STRING) {
                        rowData.add(cell.getStringCellValue());
                    } else if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
                        Date date = cell.getDateCellValue();
                        LocalDate localDate = date.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
                        String stringDate = localDate.toString();
                        rowData.add(stringDate);
                    } else {
                        rowData.add(String.valueOf(cell.getNumericCellValue()));
                    }
                }

                data.add(rowData);
            }

            LocalDate startLocalDate = LocalDate.parse(startDate);
            LocalDate endLocalDate = LocalDate.parse(endDate);

            List<List<String>> responseData = new ArrayList<>();


            for (List<String> innerList : data) {
                boolean isSkip = false;
                List<String> responseRowData = new ArrayList<>();
                for (int i = 0; i < innerList.size(); i++) {
                    // 날짜
                    if (i == 0) {
                        LocalDate date = LocalDate.parse(innerList.get(i));
                        if (date.isAfter(startLocalDate.minusDays(1)) && date.isBefore(endLocalDate.plusDays(1))) {
                            responseRowData.add(innerList.get(i));
                        } else {
                            isSkip = true;
                            break;
                        }
                    }

                    /**
                     * 대분류
                     */
                    if (i == 1) {
                        String bigCategory = innerList.get(1);
                        if (bigCategory.equals("이체") || bigCategory.equals("내계좌이체") || bigCategory.equals("미분류") ||
                                bigCategory.equals("저축") || bigCategory.equals("경조/선물") || bigCategory.equals("금융") ||
                                bigCategory.equals("금융수입") || bigCategory.equals("급여") || bigCategory.equals("기타수입") ||
                                bigCategory.equals("대출") || bigCategory.equals("카드대금") || bigCategory.equals("투자") ||
                                bigCategory.equals("현금")) {
                            isSkip = true;
                            responseRowData = new ArrayList<>();
                            break;
                        }

                        if (bigCategory.equals("교통")) {
                            responseRowData.add("교통비");
                        } else if (bigCategory.equals("문화/여가")) {
                            responseRowData.add("모임/여행/문화");
                        } else if (bigCategory.equals("뷰티/미용")) {
                            responseRowData.add("모임/여행/문화");
                        } else if (bigCategory.equals("생활")) {
                            responseRowData.add("생활용품");
                        } else if (bigCategory.equals("식비")) {
                            responseRowData.add("외식");
                        } else if (bigCategory.equals("여행/숙박")) {
                            responseRowData.add("모임/여행/문화");
                        } else if (bigCategory.equals("의료/건강")) {
                            responseRowData.add("치료");
                        } else if (bigCategory.equals("자녀/육아")) {
                            responseRowData.add("사교육+키즈카페");
                        } else if (bigCategory.equals("자동차")) {
                            responseRowData.add("주유비+차유지비");
                        } else if (bigCategory.equals("카페/간식")) {
                            responseRowData.add("커피/베이커리");
                        } else if (bigCategory.equals("패션/쇼핑")) {
                            responseRowData.add("백화점/아울렛/스타일/수동입력필요?/아니면 기타로");
                        } else {
                            responseRowData.add(bigCategory);
                        }
                    }

                    /**
                     * 소분류 (대분류 중 애매한 것 변경)
                     */
                    if (i == 2) {
                        String smallCategory = innerList.get(2);

                        if (smallCategory.equals("마트") || smallCategory.equals("편의점")) {
                            responseRowData.remove(1);
                            responseRowData.add("마트/편의점");
                        } else if (smallCategory.equals("화장품")) {
                            responseRowData.remove(1);
                            responseRowData.add("화장품");
                        } else if (smallCategory.equals("목욕")) {
                            responseRowData.remove(1);
                            responseRowData.add("모임/여행/문화");
                        } else if (smallCategory.equals("배달")) {
                            responseRowData.remove(1);
                            responseRowData.add("술/배달");
                        } else if (smallCategory.equals("휴대폰")) {
                            responseRowData.remove(1);
                            responseRowData.add("휴대폰");
                        } else if (smallCategory.equals("관리비") || smallCategory.equals("가스비")) {
                            responseRowData.remove(1);
                            responseRowData.add("관리비");
                        } else if (smallCategory.equals("인터넷")) {
                            responseRowData.remove(1);
                            responseRowData.add("인터넷/TV");
                        } else if (smallCategory.equals("아이스크림/빙수")) {
                            responseRowData.remove(1);
                            responseRowData.add("마트/편의점");
                        }
                    }

                    if (i == 3) {
                        String contents = innerList.get(i);
                        if (contents.equals("보육료)")) {
                            isSkip = true;
                            responseRowData = new ArrayList<>();
                            break;
                        }
                        responseRowData.add(contents);
                    }

                    if (i == 4) {
                        String price = innerList.get(i);
                        if (price.contains("-")) {
                            price = price.replaceAll("-", "");
                        } else {
                            price = "-" + price;
                        }
                        responseRowData.add(price);
                    }

                    if (i == 5) {
                        String card = innerList.get(i);

                        if (card.equals("성남시 아동수당  Deep Dream(체크)")) {
                            responseRowData.add("아동수당카드");
                        } else if (card.equals("올바른POINT UP+카드")) {
                            if (owner.equals("P")) {
                                responseRowData.add("광태 농협");
                            } else {
                                responseRowData.add("주희 농협");
                            }

                        } else if (card.equals("T-economy")) {
                            responseRowData.add("광태 국민");
                        } else if (card.equals("에스케이패밀리(직원) 생활밀착형 카드 ")) {
                            responseRowData.add("광태 회사 복지 카드");
                        } else if (card.equals("채움 뉴 후불 하이패스(개인)카드")) {
                            responseRowData.add("광태 농협");
                        } else if (card.equals("성남사랑상품권")) {
                            responseRowData.add("광태 농협");
                        } else if (card.equals("NH주거래우대통장") || card.equals("KB종합통장-저축예금")) {
                            responseRowData = new ArrayList<>();
                            isSkip = true;
                            break;
                        } else {
                            responseRowData.add(card);
                        }
                    }
                }

                if (!responseRowData.isEmpty()) {
                    String temp = responseRowData.get(1);
                    responseRowData.remove(1);
                    responseRowData.add(temp);
                }

                if (!isSkip) {
                    responseData.add(responseRowData);
                }
            }


            // 새 엑셀 파일을 생성하고 응답으로 반환
            Workbook responseWorkbook = new XSSFWorkbook();
            Sheet responseSheet = responseWorkbook.createSheet("Processed Data");

            Row headerRow = responseSheet.createRow(0);
            headerRow.createCell(0).setCellValue("거래일");
            headerRow.createCell(1).setCellValue("지출내용");
            headerRow.createCell(2).setCellValue("지출금액");
            headerRow.createCell(3).setCellValue("지출방법");
            headerRow.createCell(4).setCellValue("소비분류");

            int rowNo = 1;

            for (List<String> innerResponse : responseData) {
                Row row = responseSheet.createRow(rowNo++);
                for (int i = 0; i < innerResponse.size(); i++) {
                    row.createCell(i).setCellValue(innerResponse.get(i));
                }
            }

            // 엑셀 파일을 ByteArrayOutputStream에 쓰기
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            responseWorkbook.write(outputStream);
            outputStream.close();

            // 엑셀 파일을 HTTP 응답으로 보내기
            byte[] excelContent = outputStream.toByteArray();
            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
            String fileName = "processed.xlsx";
            headers.setContentDispositionFormData(fileName, fileName);
            headers.setCacheControl("must-revalidate, post-check=0, pre-check=0");

            return new ResponseEntity<>(excelContent, headers, HttpStatus.OK);
        } catch (IOException | EncryptedDocumentException e) {
            e.printStackTrace();
            return new ResponseEntity<>(new ByteArrayOutputStream().toByteArray(), HttpStatus.OK);
        }
    }
}
