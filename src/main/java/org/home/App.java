package org.home;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class App
{
    public static final int MAX_VALUE = 120;
    public static final int MULL_VALUE = 0;
    public static final int RESONABLE_VALUE = MULL_VALUE + 4;
    public static final int ROW_QNTY = 30;
    public static final int COL_QNTY = 6;
    public static final Random random = new Random();

    private static final Logger logger = LoggerFactory.getLogger(App.class);
    public static final int START_VALUE = 2;
    public static int PAGE_COUNTER = initializePageCounter();

    public static void writeToExcelFile(List<String> randomSelections, List<String> randomSelectionsAns, String filePath) {
        String filePathAns = filePath + "Ans";

        try (Workbook workbook = new HSSFWorkbook(); Workbook workbookAns = new HSSFWorkbook()) {
            String sheetName = LocalDate.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd"));
            Sheet sheet = workbook.createSheet(sheetName);
            Sheet sheetAns = workbookAns.createSheet(sheetName);

            int columnWidth = 4430;
            for (int i = 0; i < COL_QNTY; i++) {
                sheet.setColumnWidth(i, columnWidth);
                sheetAns.setColumnWidth(i, columnWidth);
            }

            CellStyle baseStyle = workbook.createCellStyle();
            Font font = workbook.createFont();
            font.setFontName("Calibri");
            font.setFontHeightInPoints((short) 11);
            baseStyle.setFont(font);
            baseStyle.setAlignment(HorizontalAlignment.CENTER);
            baseStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            CellStyle cellStyleWithBorder = workbook.createCellStyle();
            cellStyleWithBorder.cloneStyleFrom(baseStyle);
            cellStyleWithBorder.setBorderRight(BorderStyle.THIN);

            CellStyle baseStyleAns = workbookAns.createCellStyle();
            baseStyleAns.cloneStyleFrom(baseStyle);
            CellStyle cellStyleWithBorderAns = workbookAns.createCellStyle();
            cellStyleWithBorderAns.cloneStyleFrom(cellStyleWithBorder);

            int rowCount = 0;
            for (int i = 0; i < ROW_QNTY; i++) {
                Row row = sheet.createRow(rowCount);
                Row rowAns = sheetAns.createRow(rowCount++);
                row.setHeightInPoints(26.25f);
                rowAns.setHeightInPoints(26.25f);

                for (int j = 0; j < COL_QNTY; j++) {
                    int index = i * COL_QNTY + j;

                    Cell cell = row.createCell(j);
                    Cell cellAns = rowAns.createCell(j);

                    if (index < randomSelections.size()) {
                        cell.setCellValue(randomSelections.get(index));
                    } else {
                        cell.setCellValue("N/A");
                    }

                    if (index < randomSelectionsAns.size()) {
                        cellAns.setCellValue(randomSelectionsAns.get(index));
                    } else {
                        cellAns.setCellValue("N/A");
                    }

                    if (j < COL_QNTY - 1) {
                        cell.setCellStyle(cellStyleWithBorder);
                        cellAns.setCellStyle(cellStyleWithBorderAns);
                    } else {
                        cell.setCellStyle(baseStyle);
                        cellAns.setCellStyle(baseStyleAns);
                    }
                }
            }

            PrintSetup printSetup = sheet.getPrintSetup();
            printSetup.setPaperSize(PrintSetup.LETTER_PAPERSIZE);
            printSetup.setLandscape(false);

            PrintSetup printSetupAns = sheetAns.getPrintSetup();
            printSetupAns.setPaperSize(PrintSetup.LETTER_PAPERSIZE);
            printSetupAns.setLandscape(false);

            sheet.setMargin(Sheet.TopMargin, 0.3 / 2.54);
            sheet.setMargin(Sheet.BottomMargin, 0.3 / 2.54);
            sheet.setMargin(Sheet.LeftMargin, 0.3 / 2.54);
            sheet.setMargin(Sheet.RightMargin, 0.3 / 2.54);
            sheet.setMargin(Sheet.HeaderMargin, 0);
            sheet.setMargin(Sheet.FooterMargin, 0.3 / 2.54);
            sheet.getFooter().setCenter("AlexandrMathTraning-" + LocalDate.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd")));

            sheetAns.setMargin(Sheet.TopMargin, 0.3 / 2.54);
            sheetAns.setMargin(Sheet.BottomMargin, 0.3 / 2.54);
            sheetAns.setMargin(Sheet.LeftMargin, 0.3 / 2.54);
            sheetAns.setMargin(Sheet.RightMargin, 0.3 / 2.54);
            sheetAns.setMargin(Sheet.HeaderMargin, 0);
            sheetAns.setMargin(Sheet.FooterMargin, 0.3 / 2.54);
            sheetAns.getFooter().setCenter("AlexandrMathTraning-" + LocalDate.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd")));

            try (FileOutputStream fileOut = new FileOutputStream(filePath);
                 FileOutputStream fileOutAns = new FileOutputStream(filePathAns)) {
                workbook.write(fileOut);
                workbookAns.write(fileOutAns);
                System.out.println("Файлы успешно записаны: " + filePath + " и " + filePathAns);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void writeToExcelOneBook(List<String> randomSelections, List<String> randomSelectionsAns, int num) {
        String fileName = getFileName() + ".xls";

        try (Workbook workbook = new HSSFWorkbook()) {
            String sheetNameBase = LocalDate.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd"));

            // Создаем два листа в одной книге
            Sheet sheet = workbook.createSheet(sheetNameBase);
            Sheet sheetAns = workbook.createSheet(sheetNameBase + " Answers");

            // Общие настройки для обоих листов
            int columnWidth = 4430;
            for (int i = 0; i < COL_QNTY; i++) {
                sheet.setColumnWidth(i, columnWidth);
                sheetAns.setColumnWidth(i, columnWidth);
            }

            // Создаем стили для основной книги
            CellStyle baseStyle = workbook.createCellStyle();
            Font font = workbook.createFont();
            font.setFontName("Calibri");
            font.setFontHeightInPoints((short) 11);
            baseStyle.setFont(font);
            baseStyle.setAlignment(HorizontalAlignment.CENTER);
            baseStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            CellStyle cellStyleWithBorder = workbook.createCellStyle();
            cellStyleWithBorder.cloneStyleFrom(baseStyle);
            cellStyleWithBorder.setBorderRight(BorderStyle.THIN);

            // Заполняем оба листа
            fillSheet(sheet, randomSelections, baseStyle, cellStyleWithBorder);
            fillSheet(sheetAns, randomSelectionsAns, baseStyle, cellStyleWithBorder);

            // Настройки печати для основного листа
            setPrintSettings(sheet);
            setPrintSettings(sheetAns);

            // Сохраняем книгу
            try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
                workbook.write(fileOut);
                System.out.println("Файл успешно записан: " + fileName);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String getFileName() {
        return "AlexandrMathTraning-" + LocalDate.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd")) + "-" + PAGE_COUNTER;
    }

    private static void fillSheet(Sheet sheet, List<String> data, CellStyle baseStyle, CellStyle borderStyle) {
        int rowCount = 0;
        for (int i = 0; i < ROW_QNTY; i++) {
            Row row = sheet.createRow(rowCount++);
            row.setHeightInPoints(26.25f);

            for (int j = 0; j < COL_QNTY; j++) {
                int index = i * COL_QNTY + j;
                Cell cell = row.createCell(j);

                cell.setCellValue(index < data.size() ? data.get(index) : "N/A");
                cell.setCellStyle(j < COL_QNTY - 1 ? borderStyle : baseStyle);
            }
        }
    }

    private static void setPrintSettings(Sheet sheet) {
        PrintSetup printSetup = sheet.getPrintSetup();
        printSetup.setPaperSize(PrintSetup.LETTER_PAPERSIZE);
        printSetup.setLandscape(false);

        sheet.setMargin(Sheet.TopMargin, 0.3 / 2.54);
        sheet.setMargin(Sheet.BottomMargin, 0.3 / 2.54);
        sheet.setMargin(Sheet.LeftMargin, 0.3 / 2.54);
        sheet.setMargin(Sheet.RightMargin, 0.3 / 2.54);
        sheet.setMargin(Sheet.HeaderMargin, 0);
        sheet.setMargin(Sheet.FooterMargin, 0);
//        sheet.setMargin(Sheet.FooterMargin, 0.3 / 2.54);

        String headerText = getFileName();
//        String footerText = "AlexandrMathTraning-" + LocalDate.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd"));
        sheet.getHeader().setCenter(headerText);
//        sheet.getFooter().setCenter(footerText);
    }

    public static void writeToFile(List<String> randomSelections, List<String> randomSelectionsAns, String filePath) {
        String filePathAns = filePath + "Ans";

        try (BufferedWriter writer = new BufferedWriter(new FileWriter(filePath));
             BufferedWriter writerAns = new BufferedWriter(new FileWriter(filePathAns))) {

            for (int i = 0; i < ROW_QNTY; i++) {
                for (int j = 0; j < COL_QNTY; j++) {
                    int index = i * COL_QNTY + j;

                    // Запись в первый файл
                    if (index < randomSelections.size()) {
                        writer.write(randomSelections.get(index));
                    } else {
                        writer.write("N/A");
                    }

                    // Запись в второй файл
                    if (index < randomSelectionsAns.size()) {
                        writerAns.write(randomSelectionsAns.get(index));
                    } else {
                        writerAns.write("N/A");
                    }

                    // Добавляем "; " после каждого элемента, кроме последнего в строке
                    if (j < COL_QNTY - 1) {
                        writer.write("; ");
                        writerAns.write("; ");
                    }
                }
                writer.newLine(); // Перенос строки в первом файле
                writerAns.newLine(); // Перенос строки во втором файле
            }

            System.out.println("Файлы успешно записаны: " + filePath + " и " + filePathAns);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {
        Map<ArithmeticOperation, List<ArithmeticExpression>> simpleExpessionsMap = createSimpleExpessions();
        List<ArithmeticExpression> combinedList = new ArrayList<>(simpleExpessionsMap.get(ArithmeticOperation.ADD));
        combinedList.addAll(simpleExpessionsMap.get(ArithmeticOperation.SUBTRACT));
        Collections.shuffle(combinedList);

        // Выбор 144 случайных элементов
        List<ArithmeticExpression> randomSelections = new ArrayList<>();

        // Убедитесь, что количество элементов не превышает размер списка
        int selectionSize = Math.min(144, combinedList.size());

        for (int i = 0; i < selectionSize; i++) {
            // Генерируем случайный индекс и добавляем элемент в новый список
            int randomIndex = random.nextInt(combinedList.size());
            randomSelections.add(combinedList.remove(randomIndex)); // Удаляем элемент, чтобы не выбрать его снова
        }

        Map<List<ArithmeticOperation>, List<ArithmeticExpression>> complexExpressions = createComplexExpressions(simpleExpessionsMap);
        List<String> resultList = new ArrayList<>();
        List<String> resultListWithAns = new ArrayList<>();
        for (Map.Entry<List<ArithmeticOperation>, List<ArithmeticExpression>> entry : complexExpressions.entrySet()) {
            Map<Boolean, List<String>> processed = process(new ArrayList<>(entry.getValue()));
            resultList.addAll(processed.get(true));
            resultListWithAns.addAll(processed.get(false));
//            logger.info(processed.toString());
        }


        writeToFile(resultList, resultListWithAns, "output.txt");
        writeToExcelFile(resultList, resultListWithAns,"output.xls");
        writeToExcelOneBook(resultList, resultListWithAns, PAGE_COUNTER);
    }

    private static Map<Boolean, List<String>> process(List<ArithmeticExpression> sourceList) {
        Collections.shuffle(sourceList);
        int operationsQanty = ArithmeticOperation.values().length;
        int newSubsetSize = ROW_QNTY * COL_QNTY / (operationsQanty * operationsQanty);
        List<String> resultListStr = new ArrayList<>(newSubsetSize);
        List<String> resultListStrWithAns = new ArrayList<>(newSubsetSize);
        int indexToReplace = 1;
        for (int i = 0; i < newSubsetSize; i++) {
            int elementIndex = random.nextInt(sourceList.size());
//            do {
//                elementIndex = random.nextInt(sourceList.size());
//                ArithmeticExpression expression = sourceList.remove(elementIndex);
//                if ()
//            } while(true);

//            sourceList.get(elementIndex));
            String string = sourceList.remove(elementIndex).toString();
            resultListStrWithAns.add(string);
            resultListStr.add(replaceDigit(string, indexToReplace));
            indexToReplace = i % (operationsQanty * operationsQanty) + 1;
        }
        Map<Boolean, List<String>> resultMap = new HashMap<>();
        resultMap.put(true, resultListStr);
        resultMap.put(false, resultListStrWithAns);
//        Collections.shuffle(resultListStr);
        return resultMap;
    }

    private static String replaceDigit(String expression, int indexToReplace) {
        // Регулярное выражение для захвата всех частей арифметического выражения
        String regex = "(\\d{1,3})\\s([+\\-*/])\\s(\\d{1,3})\\s([+\\-*/])\\s(\\d{1,3})\\s=\\s(\\d{1,3})";
        Pattern pattern = Pattern.compile(regex);
        Matcher matcher = pattern.matcher(expression);

        if (matcher.matches()) {
            // В зависимости от индекса замены строим нужное регулярное выражение для подстановки
            switch (indexToReplace) {
                case 1:
                    return matcher.replaceFirst("___ $2 $3 $4 $5 = $6");
                case 2:
                    return matcher.replaceFirst("$1 $2 ___ $4 $5 = $6");
                case 3:
                    return matcher.replaceFirst("$1 $2 $3 $4 ___ = $6");
                case 4:
                    return matcher.replaceFirst("$1 $2 $3 $4 $5 = ___");
                default:
                    return expression;
            }
        } else {
            return expression;
        }
    }

    private static Map<List<ArithmeticOperation>, List<ArithmeticExpression>> createComplexExpressions(Map<ArithmeticOperation, List<ArithmeticExpression>> simpleExpessionsMap) {
        Map<List<ArithmeticOperation>, List<ArithmeticExpression>> complexExpessionsMap = new HashMap<>();
        for(ArithmeticOperation op : ArithmeticOperation.values()) {
            Map<ArithmeticOperation, List<ArithmeticExpression>> setOfArithmeticExpressions = new HashMap<>();
            for (Map.Entry<ArithmeticOperation, List<ArithmeticExpression>> entry : simpleExpessionsMap.entrySet()) {
                List<ArithmeticExpression> expressions = new ArrayList<>();
                ArithmeticOperation operation = entry.getKey();
                for (ArithmeticExpression exp : entry.getValue()) {
                    for (int j = RESONABLE_VALUE; j < MAX_VALUE - exp.getResult(); j++) {
                        int result = op.calculate(exp.getResult(), j);
                        if (result <= MAX_VALUE && result >= RESONABLE_VALUE) {
                            ArithmeticExpression arithmeticExpression = new ArithmeticExpression(op, new ComplexOperand(exp), new SimpleOperand(j), result);
                            expressions.add(arithmeticExpression);
                        }
                    }
                }
                complexExpessionsMap.put(Arrays.asList(operation, op), expressions);
            }
        }

        return complexExpessionsMap;
    }

    private static Map<ArithmeticOperation, List<ArithmeticExpression>> createSimpleExpessions() {
        Map<ArithmeticOperation, List<ArithmeticExpression>> setOfArithmeticExpressions = new HashMap<>();

        for(ArithmeticOperation op : ArithmeticOperation.values()) {
            List<ArithmeticExpression> expressions = new ArrayList<>();
            for (int i = START_VALUE; i < MAX_VALUE; i++) {
                for (int j = START_VALUE; j < MAX_VALUE; j++) {
                    int result = op.calculate(i, j);
                    if (result <= MAX_VALUE && result >= RESONABLE_VALUE) {
                        ArithmeticExpression arithmeticExpression = new ArithmeticExpression(op, new SimpleOperand(i), new SimpleOperand(j), result);
                        expressions.add(arithmeticExpression);
                    }
                }
            }

            Collections.shuffle(expressions);
            setOfArithmeticExpressions.put(op, expressions);
        }

        return setOfArithmeticExpressions;
    }

    // Метод для инициализации PAGE_COUNTER
    public static int initializePageCounter() {
        String currentDate = LocalDate.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd"));
        Pattern pattern = Pattern.compile(
                "^AlexandrMathTraning-" + currentDate + "-(\\d+)\\.xls$"
        );

        // Получаем текущую рабочую директорию приложения
        File directory = new File(".");
        File[] files = directory.listFiles();
        int maxCounter = 0;

        if (files != null) {
            for (File file : files) {
                Matcher matcher = pattern.matcher(file.getName());
                if (matcher.find()) {
                    int currentNumber = Integer.parseInt(matcher.group(1));
                    maxCounter = Math.max(maxCounter, currentNumber);
                }
            }
        }

        return maxCounter + 1;
    }
}
