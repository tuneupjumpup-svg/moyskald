package org.example;

import java.io.File;
import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.net.URI;
import java.net.URLEncoder;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.charset.StandardCharsets;
import java.util.*;
import java.io.BufferedReader;
import java.io.ByteArrayInputStream;
import java.io.InputStreamReader;
import java.io.FileOutputStream;
import java.util.zip.GZIPInputStream;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.DayOfWeek;
import java.time.format.DateTimeFormatter;
import java.time.temporal.TemporalAdjusters;

// Excel
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.util.CellRangeAddress;

// Email (Jakarta Mail)
import jakarta.mail.*;
import jakarta.mail.internet.*;
import jakarta.activation.DataHandler;
import jakarta.activation.FileDataSource;

public class EmployeeSalaryCalculator {

    private static final HttpClient httpClient = HttpClient.newHttpClient();
    private static final ObjectMapper mapper = new ObjectMapper();

    // Итоги по ТЕКУЩЕМУ сотруднику
    public static double totalAllServices100 = 0; // общая сумма услуг (100%) по заказам с участием сотрудника
    public static double totalWorkerServices = 0; // сумма услуг сотрудника (с учётом доли)
    public static double totalAll40 = 0;          // его 40% к выплате за неделю

    // Итоги по ВСЕМ сотрудникам для финальной таблички
    public static class WorkerTotals {
        double services100;    // общая сумма услуг (100%) по заказам с участием сотрудника
        double workerServices; // сумма услуг сотрудника
        double p40;            // 40% к выплате
    }

    // Строка для Excel
    public static class ExcelRow {
        String createdDate;    // yyyy-MM-dd
        String shippedDate;    // yyyy-MM-dd
        String orderNumber;
        String description;    // описание заказа (customerorder.description)
        String serviceName;    // название услуги или "ИТОГО по заказу"
        double serviceSum;     // сумма услуги (доля сотрудника), руб.
        double service40;      // 40% от этой суммы, руб.
        String link;           // ссылка на заказ
    }

    private static final Map<String, WorkerTotals> allTotals = new LinkedHashMap<>();
    private static final Map<String, List<ExcelRow>> excelRowsByWorker = new LinkedHashMap<>();

    // Период недели
    public static class WeekPeriod {
        public final LocalDateTime from;  // начало (включительно)
        public final LocalDateTime to;    // конец (включительно)

        public WeekPeriod(LocalDateTime from, LocalDateTime to) {
            this.from = from;
            this.to = to;
        }
    }

    // ===== SMTP-настройки (НУЖНО ЗАПОЛНИТЬ ПОД СЕБЯ) =====
    private static final String SMTP_HOST = "smtp.mail.ru";
    private static final int    SMTP_PORT = 465;
    private static final String SMTP_USER = "jumpup@mail.ru";
    private static final String SMTP_PASS  = "jR0zcDhXUvUt71Y3j67S";
    private static final String FROM_EMAIL = SMTP_USER;

    // Email-адреса сотрудников (НУЖНО ПОДСТАВИТЬ РЕАЛЬНЫЕ)
    private static final Map<String, String> workerEmails = Map.of(
            "Дмитрий", "genatsvalle@vk.com",
            "Тимур",   "umarov18@yandex.ru",
            "Даниил",  "Nova.dm96@yandex.ru",
            "Руслан",  "jumpup@mail.ru"
    );

    public static void main(String[] args) throws Exception {

        String login = "smirnova@timokhovta5";
        String password = "cascada3788932";
        String credentials = login + ":" + password;
        String authHeader = "Basic " + Base64.getEncoder()
                .encodeToString(credentials.getBytes(StandardCharsets.UTF_8));

        // 1. Считаем период пятница -> четверг
        WeekPeriod period = calcWeeklyPeriod();

        DateTimeFormatter df = DateTimeFormatter.ofPattern("yyyy-MM-dd");
        System.out.println("Неделя для расчёта:");
        System.out.println("  c  " + period.from.format(df));
        System.out.println("  по " + period.to.format(df) + " (включительно)");
        System.out.println("====================================");

        // Список сотрудников, которых считаем
        String[] workers = {"Дмитрий", "Тимур", "Даниил", "Руслан"};

        for (String workerName : workers) {
            // сбрасываем итоги для текущего сотрудника
            totalAllServices100 = 0;
            totalWorkerServices = 0;
            totalAll40 = 0;

            // считаем его неделю
            processWeeklyForWorker(workerName, authHeader, period);

            // запоминаем для финальной таблички
            WorkerTotals wt = new WorkerTotals();
            wt.services100 = totalAllServices100;
            wt.workerServices = totalWorkerServices;
            wt.p40 = totalAll40;
            allTotals.put(workerName, wt);

            // гарантируем, что для сотрудника есть список строк Excel
            excelRowsByWorker.computeIfAbsent(workerName, k -> new ArrayList<>());
        }

        // ===== Сводка по каждому сотруднику (внизу консолью) =====
        System.out.println();
        System.out.println("=========== ИТОГО ЗА НЕДЕЛЮ ПО КАЖДОМУ СОТРУДНИКУ ===========");
        for (String w : workers) {
            WorkerTotals wt = allTotals.get(w);
            if (wt == null) continue;

            System.out.println("ИТОГО ЗА НЕДЕЛЮ ДЛЯ " + w + ":");
            System.out.println("  Общая сумма услуг (100%): " + wt.services100 + " ₽");
            System.out.println("  Сумма услуг сотрудника:   " + wt.workerServices + " ₽");
            System.out.println("  40% к выплате:            " + wt.p40 + " ₽");
            System.out.println();
        }

        // ===== Финальная табличка только по 40% =====
        System.out.println("========  ИТОГО 40% ЗА НЕДЕЛЮ ПО ВСЕМ СОТРУДНИКАМ  ========");
        for (String w : workers) {
            WorkerTotals wt = allTotals.get(w);
            double p40 = (wt != null) ? wt.p40 : 0.0;
            System.out.println(w + " — " + p40 + " ₽");
        }
        System.out.println("============================================================");

        // Создаём общий Excel-отчёт
        createExcelReport(allTotals, excelRowsByWorker, period);
        System.out.println("Excel-отчёт сохранён в файле salary-report.xlsx");

        // Рассылаем личные отчёты по e-mail
        sendEmailsToWorkers(workerEmails, allTotals, excelRowsByWorker, period);
        System.out.println("Email-отчёты отправлены.");
    }

    /**
     * Считаем период: прошлая пятница 00:01 -> этот четверг 23:59
     */
    public static WeekPeriod calcWeeklyPeriod() {
        LocalDate today = LocalDate.now();

        // Последний четверг (включая сегодня, если сегодня четверг)
        LocalDate lastThursday = today.with(TemporalAdjusters.previousOrSame(DayOfWeek.THURSDAY));

        // Пятница за 6 дней до него
        LocalDate lastFriday = lastThursday.minusDays(6);

        LocalDateTime from = lastFriday.atTime(0, 1);     // пятница 00:01
        LocalDateTime to   = lastThursday.atTime(23, 59); // четверг 23:59

        return new WeekPeriod(from, to);
    }

    /**
     * Обрабатываем все заказы за неделю для одного сотрудника.
     */
    public static void processWeeklyForWorker(String workerName,
                                              String authHeader,
                                              WeekPeriod period) throws Exception {

        System.out.println();
        System.out.println("===== ОТЧЁТ ПО СОТРУДНИКУ: " + workerName + " =====");

        // фильтр по updated за период
        DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
        String fromStr = period.from.format(dtf);
        String toStr   = period.to.format(dtf);

        String filterExpr = "updated>=" + fromStr + ";updated<=" + toStr;
        String encodedFilter = URLEncoder.encode(filterExpr, StandardCharsets.UTF_8);

        String url = "https://api.moysklad.ru/api/remap/1.2/entity/customerorder"
                + "?limit=50&order=updated,desc&filter=" + encodedFilter;

        int page = 1;

        while (url != null && !url.isEmpty()) {
            System.out.println("---- Страница " + page + " ----");
            System.out.println("Загружаю: " + url);

            HttpRequest request = HttpRequest.newBuilder()
                    .uri(URI.create(url))
                    .header("Authorization", authHeader)
                    .header("Accept", "application/json;charset=utf-8")
                    .header("Content-Type", "application/json")
                    .header("Lognex-Pretty-Print-Json", "true")
                    .header("Accept-Encoding", "gzip")
                    .GET()
                    .build();

            HttpResponse<byte[]> response = httpClient.send(request, HttpResponse.BodyHandlers.ofByteArray());
            String bodyText = unpack(response);
            System.out.println("Страница загружена.");

            JsonNode root = mapper.readTree(bodyText);
            JsonNode rows = root.path("rows");
            if (!rows.isArray() || rows.size() == 0) {
                break;
            }

            for (JsonNode order : rows) {
                processSingleOrderForWorker(order, workerName, authHeader, period);
            }

            String nextHref = root.path("meta").path("nextHref").asText("");
            if (nextHref == null || nextHref.isBlank()) {
                url = null;
            } else {
                url = nextHref;
                page++;
            }
        }

        System.out.println("Обход заказов завершён.");
        System.out.println();
        System.out.println("ИТОГО ЗА НЕДЕЛЮ ДЛЯ " + workerName + ":");
        System.out.println("  Общая сумма услуг (100%): " + totalAllServices100 + " ₽");
        System.out.println("  Сумма услуг сотрудника:   " + totalWorkerServices + " ₽");
        System.out.println("  40% к выплате:            " + totalAll40 + " ₽");
        System.out.println("========================================");
    }

    /**
     * Обработка одного заказа: фильтры + расчёт по конкретному сотруднику.
     * Обязательно должна быть ОТГРУЗКА, и именно её дата попадает в период.
     */
    public static void processSingleOrderForWorker(JsonNode order,
                                                   String workerName,
                                                   String authHeader,
                                                   WeekPeriod period) throws Exception {
        String orderName = order.path("name").asText("");
        String description = order.path("description").asText(""); // ✅ берём описание

        // 1) Сумма заказа (100%). Нулевые/отрицательные — сразу мимо.
        double sumMinor = order.path("sum").asDouble();
        if (sumMinor <= 0) return;

        // 2) Ищем дату ОТГРУЗКИ (обязательно!)
        LocalDateTime shipMoment = null;
        JsonNode demands = order.path("demands");
        if (demands.isArray() && demands.size() > 0) {
            JsonNode firstDemand = demands.get(0);
            String demandHref = firstDemand.path("meta").path("href").asText("");
            if (!demandHref.isEmpty()) {
                HttpRequest demReq = HttpRequest.newBuilder()
                        .uri(URI.create(demandHref))
                        .header("Authorization", authHeader)
                        .header("Accept", "application/json;charset=utf-8")
                        .header("Content-Type", "application/json")
                        .header("Lognex-Pretty-Print-Json", "true")
                        .header("Accept-Encoding", "gzip")
                        .GET()
                        .build();

                HttpResponse<byte[]> demResp = httpClient.send(demReq, HttpResponse.BodyHandlers.ofByteArray());
                String demBody = unpack(demResp);
                JsonNode demRoot = mapper.readTree(demBody);
                String momentStr = demRoot.path("moment").asText("");
                shipMoment = parseMoment(momentStr);
            }
        }

        // Если отгрузки нет — заказ в зарплату НЕ идёт
        if (shipMoment == null) {
            return;
        }

        // Проверяем, что дата ОТГРУЗКИ попадает в нашу неделю
        if (shipMoment.isBefore(period.from) || shipMoment.isAfter(period.to)) {
            return;
        }

        // 3) Исполнители
        String worker1 = null;
        String worker2 = null;

        JsonNode attributes = order.path("attributes");
        if (attributes.isArray()) {
            for (JsonNode attr : attributes) {
                String attrName = attr.path("name").asText();
                if ("Студия Исполнитель_1".equals(attrName)) {
                    worker1 = attr.path("value").path("name").asText(null);
                }
                if ("Студия Исполнитель_2".equals(attrName)) {
                    worker2 = attr.path("value").path("name").asText(null);
                }
            }
        }

        int executorsCount = 0;
        if (worker1 != null && !worker1.isBlank()) executorsCount++;
        if (worker2 != null && !worker2.isBlank()) executorsCount++;

        boolean ourWorkerInOrder =
                workerName.equals(worker1) || workerName.equals(worker2);

        // если наш сотрудник не участвует или нет ни одного исполнителя — выходим
        if (!ourWorkerInOrder || executorsCount == 0) {
            return;
        }

        // 4) Даты и ссылка
        String createdStr = order.path("moment").asText("");
        LocalDateTime createdMoment = parseMoment(createdStr);
        String createdDate = createdMoment != null ? createdMoment.toLocalDate().toString() : "";
        String shippedDate = shipMoment.toLocalDate().toString();

        String link = order.path("meta").path("uuidHref").asText("");

        // список строк Excel для этого сотрудника
        List<ExcelRow> excelRows = excelRowsByWorker
                .computeIfAbsent(workerName, k -> new ArrayList<>());

        // 5) Позиции — услуги с флагом "Учёт в зарплате"
        String positionsHref = order.path("positions").path("meta").path("href").asText("");
        if (positionsHref == null || positionsHref.isBlank()) {
            return;
        }

        String positionsUrl = positionsHref + "?expand=assortment";

        HttpRequest posRequest = HttpRequest.newBuilder()
                .uri(URI.create(positionsUrl))
                .header("Authorization", authHeader)
                .header("Accept", "application/json;charset=utf-8")
                .header("Content-Type", "application/json")
                .header("Lognex-Pretty-Print-Json", "true")
                .header("Accept-Encoding", "gzip")
                .GET()
                .build();

        HttpResponse<byte[]> posResponse = httpClient.send(posRequest, HttpResponse.BodyHandlers.ofByteArray());
        String posText = unpack(posResponse);

        JsonNode posRoot = mapper.readTree(posText);
        JsonNode positions = posRoot.path("rows");

        double totalServicesMinor     = 0.0; // 100% сумма услуг (копейки)
        double workerServicesOrderRub = 0.0; // сумма услуг сотрудника по заказу
        double worker40OrderRub       = 0.0; // его 40% по заказу

        if (positions.isArray()) {
            for (JsonNode pos : positions) {
                String type = pos.path("assortment").path("meta").path("type").asText();
                if (!"service".equalsIgnoreCase(type)) continue;

                JsonNode assortment = pos.path("assortment");
                boolean includeInSalary = false;
                JsonNode assAttrs = assortment.path("attributes");
                if (assAttrs.isArray()) {
                    for (JsonNode a : assAttrs) {
                        String aName = a.path("name").asText();
                        if ("Учёт в зарплате".equals(aName)) {
                            includeInSalary = a.path("value").asBoolean(false);
                        }
                    }
                }
                if (!includeInSalary) continue;

                double priceMinor   = pos.path("price").asDouble();
                double qty          = pos.path("quantity").asDouble();
                double lineSumMinor = priceMinor * qty;

                totalServicesMinor += lineSumMinor;

                double lineRub = lineSumMinor / 100.0;

                double shareFactor      = 1.0 / executorsCount;
                double workerServiceRub = lineRub * shareFactor;
                double worker40Rub      = workerServiceRub * 0.40;

                workerServicesOrderRub += workerServiceRub;
                worker40OrderRub       += worker40Rub;

                // строка по услуге
                ExcelRow er = new ExcelRow();
                er.createdDate = createdDate;
                er.shippedDate = shippedDate;
                er.orderNumber = orderName;
                er.description = description; // ✅ кладём описание в строку
                er.serviceName = assortment.path("name").asText("");
                er.serviceSum  = workerServiceRub;
                er.service40   = worker40Rub;
                er.link        = link;
                excelRows.add(er);
            }
        }

        // если нет услуг с "Учёт в зарплате" — заказ пропускаем
        if (totalServicesMinor == 0.0 || workerServicesOrderRub == 0.0) {
            return;
        }

        double totalServicesRub = totalServicesMinor / 100.0;

        // строка "ИТОГО по заказу"
        ExcelRow totalRow = new ExcelRow();
        totalRow.createdDate = createdDate;
        totalRow.shippedDate = shippedDate;
        totalRow.orderNumber = orderName;
        totalRow.description = description; // ✅ чтобы в итого тоже было
        totalRow.serviceName = "ИТОГО по заказу";
        totalRow.serviceSum  = workerServicesOrderRub;
        totalRow.service40   = worker40OrderRub;
        totalRow.link        = link;
        excelRows.add(totalRow);

        // накапливаем итоги по сотруднику
        totalAllServices100 += totalServicesRub;
        totalWorkerServices += workerServicesOrderRub;
        totalAll40          += worker40OrderRub;
    }

    private static String resolveFreeFileName(String baseName) {
        File f = new File(baseName);

        if (!f.exists()) {
            return baseName;
        }

        String prefix;
        String suffix;

        int dot = baseName.lastIndexOf('.');
        if (dot == -1) {
            prefix = baseName;
            suffix = "";
        } else {
            prefix = baseName.substring(0, dot);
            suffix = baseName.substring(dot); // включая точку
        }

        int counter = 1;
        String name;
        do {
            name = prefix + " (" + counter + ")" + suffix;
            f = new File(name);
            counter++;
        } while (f.exists());

        return name;
    }

    /**
     * Распаковка возможного gzip-ответа
     */
    public static String unpack(HttpResponse<byte[]> response) throws Exception {
        byte[] bytes = response.body();

        Optional<String> enc = response.headers().firstValue("Content-Encoding");
        if (enc.isPresent() && enc.get().contains("gzip")) {
            try (GZIPInputStream gis = new GZIPInputStream(new ByteArrayInputStream(bytes));
                 InputStreamReader isr = new InputStreamReader(gis, StandardCharsets.UTF_8);
                 BufferedReader br = new BufferedReader(isr)) {

                StringBuilder sb = new StringBuilder();
                String line;
                while ((line = br.readLine()) != null) {
                    sb.append(line);
                }
                return sb.toString();
            }
        }
        return new String(bytes, StandardCharsets.UTF_8);
    }

    /**
     * Парсер дат вида "2025-11-07 18:05:00.000"
     */
    public static LocalDateTime parseMoment(String momentStr) {
        if (momentStr == null || momentStr.isBlank()) {
            return null;
        }
        try {
            String withoutMillis = momentStr.split("\\.")[0]; // "2025-11-07 18:05:00"
            DateTimeFormatter fmt = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
            return LocalDateTime.parse(withoutMillis, fmt);
        } catch (Exception e) {
            System.out.println("Не удалось распарсить дату: " + momentStr);
            return null;
        }
    }

    /**
     * Создание общего Excel-отчёта по всем сотрудникам.
     */
    public static void createExcelReport(Map<String, WorkerTotals> allTotals,
                                         Map<String, List<ExcelRow>> rowsByWorker,
                                         WeekPeriod period) throws Exception {

        Workbook wb = new XSSFWorkbook();

        // стили
        CellStyle headerStyle = wb.createCellStyle();
        Font headerFont = wb.createFont();
        headerFont.setBold(true);
        headerStyle.setFont(headerFont);

        // стиль ссылки
        CellStyle linkStyle = wb.createCellStyle();
        Font linkFont = wb.createFont();
        linkFont.setUnderline(Font.U_SINGLE);
        linkFont.setColor(IndexedColors.BLUE.getIndex());
        linkStyle.setFont(linkFont);

        // стиль чисел с разделителем тысяч
        DataFormat dataFormat = wb.createDataFormat();
        short numFmt = dataFormat.getFormat("# ##0");
        CellStyle numberStyle = wb.createCellStyle();
        numberStyle.setDataFormat(numFmt);

        // ✅ стиль переноса строк для описания
        CellStyle wrapStyle = wb.createCellStyle();
        wrapStyle.setWrapText(true);

        CreationHelper creationHelper = wb.getCreationHelper();

        String periodText = "Неделя для расчёта: c " +
                period.from.toLocalDate() +
                " по " +
                period.to.toLocalDate() +
                " (включительно)";

        String[] headers = {
                "Дата создания",
                "Дата отгрузки",
                "Номер заказа",
                "Описание",
                "Услуга",
                "Сумма услуги",
                "40%",
                "Ссылка"
        };

        // ---------- ЛИСТЫ ПО СОТРУДНИКАМ ----------
        for (Map.Entry<String, List<ExcelRow>> entry : rowsByWorker.entrySet()) {
            String workerName = entry.getKey();
            List<ExcelRow> rows = entry.getValue();

            Sheet sheet = wb.createSheet(workerName);

            int rowIdx = 0;

            // строка с периодом
            Row periodRow = sheet.createRow(rowIdx++);
            Cell pCell = periodRow.createCell(0);
            pCell.setCellValue(periodText);
            pCell.setCellStyle(headerStyle);
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, headers.length - 1));

            rowIdx++; // пустая строка

            // шапка
            Row header = sheet.createRow(rowIdx++);
            for (int i = 0; i < headers.length; i++) {
                Cell c = header.createCell(i);
                c.setCellValue(headers[i]);
                c.setCellStyle(headerStyle);
            }

            int dataStartRow = rowIdx; // первая строка данных

            // строки по услугам и итогам по заказам
            for (ExcelRow er : rows) {
                Row r = sheet.createRow(rowIdx++);

                r.createCell(0).setCellValue(er.createdDate != null ? er.createdDate : "");
                r.createCell(1).setCellValue(er.shippedDate != null ? er.shippedDate : "");
                r.createCell(2).setCellValue(er.orderNumber != null ? er.orderNumber : "");

                Cell descCell = r.createCell(3);
                descCell.setCellValue(er.description != null ? er.description : "");
                descCell.setCellStyle(wrapStyle);

                r.createCell(4).setCellValue(er.serviceName != null ? er.serviceName : "");

                Cell sumCell = r.createCell(5);
                sumCell.setCellValue(er.serviceSum);
                sumCell.setCellStyle(numberStyle);

                Cell p40Cell = r.createCell(6);
                p40Cell.setCellValue(er.service40);
                p40Cell.setCellStyle(numberStyle);

                Cell linkCell = r.createCell(7);
                if (er.link != null && !er.link.isBlank()) {
                    linkCell.setCellValue("Открыть");
                    Hyperlink hyperlink = creationHelper.createHyperlink(HyperlinkType.URL);
                    hyperlink.setAddress(er.link);
                    linkCell.setHyperlink(hyperlink);
                    linkCell.setCellStyle(linkStyle);
                } else {
                    linkCell.setCellValue("");
                }
            }

            // автоширина
            for (int i = 0; i < headers.length; i++) {
                sheet.autoSizeColumn(i, true);
                int width = sheet.getColumnWidth(i);
                sheet.setColumnWidth(i, width + 512); // небольшой запас
            }

            // объединяем ячейки по заказам (A,B,C,H)
            mergeOrderBlocks(sheet, rows, dataStartRow);

            // в конце — строка ИТОГО по сотруднику
            WorkerTotals wt = allTotals.get(workerName);
            if (wt != null) {
                sheet.createRow(rowIdx++); // пустая строка
                Row totalRow = sheet.createRow(rowIdx++);

                Cell c0 = totalRow.createCell(0);
                c0.setCellValue("ИТОГО ПО СОТРУДНИКУ " + workerName);
                c0.setCellStyle(headerStyle);

                Cell totalServicesCell = totalRow.createCell(5); // ✅ было 4
                totalServicesCell.setCellValue(wt.workerServices);
                totalServicesCell.setCellStyle(numberStyle);

                Cell total40Cell = totalRow.createCell(6); // ✅ было 5
                total40Cell.setCellValue(wt.p40);
                total40Cell.setCellStyle(numberStyle);
            }

            // фильтр по шапке
            int headerRowIndex = 2; // период (0), пустая (1), шапка (2)
            sheet.setAutoFilter(new CellRangeAddress(headerRowIndex,
                    rowIdx - 1, 0, headers.length - 1));
        }

        // ---------- ЛИСТ СВОДКА ----------
        Sheet summary = wb.createSheet("Сводка");
        int rowIdx = 0;

        Row periodRow = summary.createRow(rowIdx++);
        Cell pCell = periodRow.createCell(0);
        pCell.setCellValue(periodText);
        pCell.setCellStyle(headerStyle);
        summary.addMergedRegion(new CellRangeAddress(0, 0, 0, 2));

        rowIdx++; // пустая строка

        Row h = summary.createRow(rowIdx++);
        String[] sumHeaders = {"Сотрудник", "Сумма услуг", "40% за неделю"};
        for (int i = 0; i < sumHeaders.length; i++) {
            Cell c = h.createCell(i);
            c.setCellValue(sumHeaders[i]);
            c.setCellStyle(headerStyle);
        }

        double grandServices = 0.0;
        double grand40 = 0.0;

        for (Map.Entry<String, WorkerTotals> e : allTotals.entrySet()) {
            String workerName = e.getKey();
            WorkerTotals wt = e.getValue();

            Row r = summary.createRow(rowIdx++);
            r.createCell(0).setCellValue(workerName);

            Cell sCell = r.createCell(1);
            sCell.setCellValue(wt.workerServices);
            sCell.setCellStyle(numberStyle);

            Cell pCell40 = r.createCell(2);
            pCell40.setCellValue(wt.p40);
            pCell40.setCellStyle(numberStyle);

            grandServices += wt.workerServices;
            grand40 += wt.p40;
        }

        Row totalSummaryRow = summary.createRow(rowIdx++);
        Cell t0 = totalSummaryRow.createCell(0);
        t0.setCellValue("ИТОГО ВСЕГО");
        t0.setCellStyle(headerStyle);

        Cell tsCell = totalSummaryRow.createCell(1);
        tsCell.setCellValue(grandServices);
        tsCell.setCellStyle(numberStyle);

        Cell tpCell = totalSummaryRow.createCell(2);
        tpCell.setCellValue(grand40);
        tpCell.setCellStyle(numberStyle);

        for (int i = 0; i < sumHeaders.length; i++) {
            summary.autoSizeColumn(i, true);
            int width = summary.getColumnWidth(i);
            summary.setColumnWidth(i, width + 512);
        }

        try (FileOutputStream fos = new FileOutputStream("salary-report.xlsx")) {
            wb.write(fos);
        }
        wb.close();
    }

    /**
     * Объединение ячеек по заказам (для колонок A,B,C,H).
     */
    private static void mergeOrderBlocks(Sheet sheet, List<ExcelRow> rows, int dataStartRow) {
        int currentRow = dataStartRow;
        for (int i = 0; i < rows.size(); ) {
            ExcelRow first = rows.get(i);
            String order = first.orderNumber;
            if (order == null || order.isBlank()) {
                currentRow++;
                i++;
                continue;
            }

            int groupSize = 1;
            int j = i + 1;
            while (j < rows.size()
                    && Objects.equals(order, rows.get(j).orderNumber)) {
                groupSize++;
                j++;
            }

            if (groupSize > 1) {
                int fromRow = currentRow;
                int toRow = currentRow + groupSize - 1;

                int[] colsToMerge = {0, 1, 2, 7}; // ✅ даты, номер, ссылка (было 6)
                for (int col : colsToMerge) {
                    sheet.addMergedRegion(new CellRangeAddress(fromRow, toRow, col, col));
                }
            }

            currentRow += groupSize;
            i += groupSize;
        }
    }

    // ============================ EMAIL-БЛОК ============================

    /**
     * Разослать личные отчёты всем сотрудникам.
     */
    public static void sendEmailsToWorkers(Map<String, String> emails,
                                           Map<String, WorkerTotals> totals,
                                           Map<String, List<ExcelRow>> rowsByWorker,
                                           WeekPeriod period) throws Exception {

        for (Map.Entry<String, String> e : emails.entrySet()) {
            String workerName = e.getKey();
            String email = e.getValue();

            List<ExcelRow> rows = rowsByWorker.get(workerName);
            WorkerTotals wt = totals.get(workerName);

            if (rows == null || rows.isEmpty() || wt == null) {
                System.out.println("Нет данных для " + workerName + ", email не отправляем.");
                continue;
            }

            String fileName = "salary-" + workerName + "-" +
                    period.from.toLocalDate() + "_" + period.to.toLocalDate() + ".xlsx";

            createPersonalExcelForWorker(workerName, rows, wt, period, fileName);

            sendEmailWithAttachment(
                    email,
                    "Отчёт по студии за неделю " +
                            period.from.toLocalDate() + " - " + period.to.toLocalDate(),
                    buildEmailBody(workerName, wt, period),
                    fileName
            );
        }
    }

    /**
     * Личный Excel-файл для одного сотрудника.
     */
    public static void createPersonalExcelForWorker(String workerName,
                                                    List<ExcelRow> rows,
                                                    WorkerTotals wt,
                                                    WeekPeriod period,
                                                    String fileName) throws Exception {

        Workbook wb = new XSSFWorkbook();

        // стили
        CellStyle headerStyle = wb.createCellStyle();
        Font headerFont = wb.createFont();
        headerFont.setBold(true);
        headerStyle.setFont(headerFont);

        DataFormat dataFormat = wb.createDataFormat();
        short numFmt = dataFormat.getFormat("# ##0");
        CellStyle numberStyle = wb.createCellStyle();
        numberStyle.setDataFormat(numFmt);

        // ✅ стиль переноса строк для описания
        CellStyle wrapStyle = wb.createCellStyle();
        wrapStyle.setWrapText(true);

        CreationHelper helper = wb.getCreationHelper();

        CellStyle linkStyle = wb.createCellStyle();
        Font linkFont = wb.createFont();
        linkFont.setUnderline(Font.U_SINGLE);
        linkFont.setColor(IndexedColors.BLUE.getIndex());
        linkStyle.setFont(linkFont);

        String periodText = "Неделя для расчёта: c " +
                period.from.toLocalDate() +
                " по " +
                period.to.toLocalDate() +
                " (включительно)";

        String[] headers = {
                "Дата создания",
                "Дата отгрузки",
                "Номер заказа",
                "Описание",
                "Услуга",
                "Сумма услуги",
                "40%",
                "Ссылка"
        };

        Sheet sheet = wb.createSheet(workerName);

        int rowIdx = 0;

        // период
        Row periodRow = sheet.createRow(rowIdx++);
        Cell pCell = periodRow.createCell(0);
        pCell.setCellValue(periodText);
        pCell.setCellStyle(headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, headers.length - 1));

        rowIdx++; // пустая строка

        // шапка
        Row header = sheet.createRow(rowIdx++);
        for (int i = 0; i < headers.length; i++) {
            Cell c = header.createCell(i);
            c.setCellValue(headers[i]);
            c.setCellStyle(headerStyle);
        }

        int dataStartRow = rowIdx;

        // данные
        for (ExcelRow er : rows) {
            Row r = sheet.createRow(rowIdx++);

            r.createCell(0).setCellValue(er.createdDate != null ? er.createdDate : "");
            r.createCell(1).setCellValue(er.shippedDate != null ? er.shippedDate : "");
            r.createCell(2).setCellValue(er.orderNumber != null ? er.orderNumber : "");

            Cell descCell = r.createCell(3);
            descCell.setCellValue(er.description != null ? er.description : "");
            descCell.setCellStyle(wrapStyle);

            r.createCell(4).setCellValue(er.serviceName != null ? er.serviceName : "");

            Cell sCell = r.createCell(5);
            sCell.setCellValue(er.serviceSum);
            sCell.setCellStyle(numberStyle);

            Cell p40Cell = r.createCell(6);
            p40Cell.setCellValue(er.service40);
            p40Cell.setCellStyle(numberStyle);

            Cell linkCell = r.createCell(7);
            if (er.link != null && !er.link.isBlank()) {
                linkCell.setCellValue("Открыть");
                Hyperlink link = helper.createHyperlink(HyperlinkType.URL);
                link.setAddress(er.link);
                linkCell.setHyperlink(link);
                linkCell.setCellStyle(linkStyle);
            } else {
                linkCell.setCellValue("");
            }
        }

        // автоширина
        for (int i = 0; i < headers.length; i++) {
            sheet.autoSizeColumn(i, true);
            int width = sheet.getColumnWidth(i);
            sheet.setColumnWidth(i, width + 512);
        }

        // объединение по заказам
        mergeOrderBlocks(sheet, rows, dataStartRow);

        // итого по сотруднику
        sheet.createRow(rowIdx++);
        Row totalRow = sheet.createRow(rowIdx++);

        Cell c0 = totalRow.createCell(0);
        c0.setCellValue("ИТОГО ПО СОТРУДНИКУ " + workerName);
        c0.setCellStyle(headerStyle);

        Cell totalServCell = totalRow.createCell(5); // ✅ было 4
        totalServCell.setCellValue(wt.workerServices);
        totalServCell.setCellStyle(numberStyle);

        Cell total40Cell = totalRow.createCell(6); // ✅ было 5
        total40Cell.setCellValue(wt.p40);
        total40Cell.setCellStyle(numberStyle);

        // фильтр
        sheet.setAutoFilter(new CellRangeAddress(2, rowIdx - 1, 0, headers.length - 1));

        try (FileOutputStream fos = new FileOutputStream(fileName)) {
            wb.write(fos);
        }
        wb.close();
    }

    /**
     * Текст письма сотруднику.
     */
    public static String buildEmailBody(String workerName,
                                        WorkerTotals wt,
                                        WeekPeriod period) {
        return "Привет, " + workerName + "!\n\n" +
                "Отчёт по студии за период:\n" +
                "c " + period.from.toLocalDate() + " по " + period.to.toLocalDate() + " (включительно).\n\n" +
                "Сумма услуг по твоим заказам: " + String.format("%,.0f", wt.workerServices).replace(',', ' ') + " ₽\n" +
                "Твои 40% за неделю:          " + String.format("%,.0f", wt.p40).replace(',', ' ') + " ₽\n\n" +
                "В приложении — подробный отчёт по заказам и услугам.\n\n" +
                "Если что-то не сходится — напиши мне.";
    }

    /**
     * Отправка письма с вложением.
     */
    public static void sendEmailWithAttachment(String toEmail,
                                               String subject,
                                               String body,
                                               String attachmentPath) throws Exception {

        Properties props = new Properties();
        props.put("mail.smtp.auth", "true");
        props.put("mail.smtp.ssl.enable", "true");
        props.put("mail.smtp.host", SMTP_HOST);
        props.put("mail.smtp.port", String.valueOf(SMTP_PORT));
        props.put("mail.debug", "false");

        Session session = Session.getInstance(props, new jakarta.mail.Authenticator() {
            @Override
            protected PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication(SMTP_USER, SMTP_PASS);
            }
        });

        Message message = new MimeMessage(session);
        message.setFrom(new InternetAddress(FROM_EMAIL));
        message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(toEmail));
        message.setSubject(subject);

        // тело письма
        MimeBodyPart textPart = new MimeBodyPart();
        textPart.setText(body, "UTF-8");

        // вложение
        MimeBodyPart attachPart = new MimeBodyPart();
        FileDataSource fds = new FileDataSource(attachmentPath);
        attachPart.setDataHandler(new DataHandler(fds));
        attachPart.setFileName(fds.getName());

        Multipart multipart = new MimeMultipart();
        multipart.addBodyPart(textPart);
        multipart.addBodyPart(attachPart);

        message.setContent(multipart);

        Transport.send(message);
        System.out.println("Письмо отправлено: " + toEmail);
    }
}