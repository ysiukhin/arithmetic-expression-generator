package org.home;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedWriter;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
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

    public static void writeToExcelFile(List<String> randomSelections, String filePath) {
        try (Workbook workbook = new HSSFWorkbook()) {
            String sheetName = LocalDate.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd"));
            Sheet sheet = workbook.createSheet(sheetName);

            // Установка ширины всех столбцов в 15.3 символа
//            int columnWidth = (int)(16.57 * 256);
            int columnWidth = 4430;
            for (int i = 0; i < COL_QNTY; i++) {
                sheet.setColumnWidth(i, columnWidth);
            }

            // Создание стиля ячеек
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

            int rowCount = 0;
            for (int i = 0; i < ROW_QNTY; i++) {
                Row row = sheet.createRow(rowCount++);
                row.setHeightInPoints(26.25f);

                for (int j = 0; j < COL_QNTY; j++) {
                    Cell cell = row.createCell(j);
                    if (i * COL_QNTY + j < randomSelections.size()) {
                        cell.setCellValue(randomSelections.get(i * COL_QNTY + j));
                    } else {
                        cell.setCellValue("N/A");
                    }

                    if (j < COL_QNTY - 1) {
                        cell.setCellStyle(cellStyleWithBorder);
                    } else {
                        cell.setCellStyle(baseStyle);
                    }
                }
            }

            // Настройка параметров печати
            PrintSetup printSetup = sheet.getPrintSetup();
            printSetup.setPaperSize(PrintSetup.LETTER_PAPERSIZE); // размер бумаги - Letter
            printSetup.setLandscape(false); // для печати в альбомной ориентации

            // Установка полей (отступов) страницы
            sheet.setMargin(Sheet.TopMargin, 0.3 / 2.54); // переводим см в дюймы
            sheet.setMargin(Sheet.BottomMargin, 0.3 / 2.54);
            sheet.setMargin(Sheet.LeftMargin, 0.3 / 2.54);
            sheet.setMargin(Sheet.RightMargin, 0.3 / 2.54);
            sheet.setMargin(Sheet.HeaderMargin, 0.8 / 2.54);
            sheet.setMargin(Sheet.FooterMargin, 0.8 / 2.54);

            try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
                workbook.write(fileOut);
                System.out.println("Файл успешно записан: " + filePath);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void writeToFile(List<String> randomSelections, String filePath) {
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(filePath))) {
            for (int i = 0; i < ROW_QNTY; i++) {
                for (int j = 0; j < COL_QNTY; j++) {
                    // Проверка на наличие доступных элементов
                    if (i * COL_QNTY + j < randomSelections.size()) {
                        writer.write(randomSelections.get(i * COL_QNTY + j).toString());
                    } else {
                        writer.write("N/A"); // Если элементов не хватает, записываем "N/A"
                    }

                    // Добавляем ";" после каждого элемента, кроме последнего в строке
                    if (j < COL_QNTY - 1) {
                        writer.write("; ");
                    }
                }
                writer.newLine(); // Перенос на новую строку после каждой строки
            }
            System.out.println("Файл успешно записан: " + filePath);
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
        for (Map.Entry<List<ArithmeticOperation>, List<ArithmeticExpression>> entry : complexExpressions.entrySet()) {
//            List<ArithmeticExpression> processed = process(new ArrayList<>(entry.getValue()));
            resultList.addAll(process(new ArrayList<>(entry.getValue())));
//            System.out.println(processed);
        }


        writeToFile(resultList, "output.txt");
        writeToExcelFile(resultList, "output.xls");
    }

    private static List<String> process(List<ArithmeticExpression> sourceList) {
        Collections.shuffle(sourceList);
        int operationsQanty = ArithmeticOperation.values().length;
        int newSubsetSize = ROW_QNTY * COL_QNTY / (operationsQanty * operationsQanty);
        List<String> resultListStr = new ArrayList<>(newSubsetSize);
        int indexToReplace = 1;
        for (int i = 0; i < newSubsetSize; i++) {
            int elementIndex = random.nextInt(sourceList.size());
//            do {
//                elementIndex = random.nextInt(sourceList.size());
//                ArithmeticExpression expression = sourceList.remove(elementIndex);
//                if ()
//            } while(true);

//            sourceList.get(elementIndex));
            resultListStr.add(replaceDigit(sourceList.remove(elementIndex).toString(), indexToReplace));
            indexToReplace = i % (operationsQanty * operationsQanty) + 1;
        }
//        Collections.shuffle(resultListStr);
        return resultListStr;
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
}
