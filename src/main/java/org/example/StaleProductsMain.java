package org.example;
import java.time.YearMonth;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import jakarta.mail.*;
import jakarta.mail.internet.InternetAddress;
import jakarta.mail.internet.MimeBodyPart;
import jakarta.mail.internet.MimeMessage;
import jakarta.mail.internet.MimeMultipart;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.net.URI;
import java.net.URISyntaxException;
import java.net.URLEncoder;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.charset.StandardCharsets;
import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;
import java.time.LocalDate;
import java.util.*;
import java.util.zip.GZIPInputStream;

public class StaleProductsMain {
    // href группы контрагентов "Розничный покупатель"
    private static final String RETAIL_GROUP_HREF =
            "https://api.moysklad.ru/api/remap/1.2/entity/group/13bc05c6-50c6-11ea-0a80-03b60000adab";

    // ====== Авторизация MoySklad ======
    private static final String LOGIN = "smirnova@timokhovta5";
    private static final String PASSWORD = "cascada3788932";
    private static final String BASE_URL = "https://api.moysklad.ru/api/remap/1.2";

    // ====== SMTP (Mail.ru) ======
    private static final String SMTP_HOST = "smtp.mail.ru";
    private static final int SMTP_PORT = 465; // SSL
    private static final String SMTP_USER = "jumpup@mail.ru";
    private static final String SMTP_PASSWORD = "jR0zcDhXUvUt71Y3j67S";
    private static final String EMAIL_FROM = SMTP_USER;
    private static final String EMAIL_TO = "jumpup@mail.ru";

    // ====== Продавцы (employee) ======
    private static final String SHELEST_ID = "0798af0d-5316-11ea-0a80-007e0017da6c";
    private static final String MIRZ_ID = "5fbba0d2-95bb-11ea-0a80-01930003a188";
    private static final String SHELEST_NAME = "Шелест";
    private static final String MIRZ_NAME = "Мирзоянов";

    // ====== Группа розничных клиентов ======
    private static final String RETAIL_GROUP_NAME = "Розничный покупатель";

    // Оборот за период по рознице (retaildemand) по ВСЕМ складам (но только по розничной группе)
    private static double totalTurnover180 = 0.0;
    private static final double MONTHS_IN_PERIOD = 6.0;

    // Динамическая шкала доли плана по неликвиду
    private static final double MIN_STALE_PLAN_SHARE = 0.05; // 5%
    private static final double MAX_STALE_PLAN_SHARE = 0.35; // 35%

    private static final ObjectMapper MAPPER = new ObjectMapper();
    private static final HttpClient httpClient = HttpClient.newBuilder()
            .followRedirects(HttpClient.Redirect.ALWAYS)
            .build();

    // ====== Параметры анализа ======
    private static final int SALES_LOOKBACK_DAYS = 180;   // период анализа продаж
    private static final int MIN_AGE_DAYS = 300;          // товар старше N дней — кандидат в неликвид
    private static final double FAST_TURNOVER_DAYS = 90;  // быстрый оборот — исключаем даже если старый
    private static final double RETAIL_BUDGET = 300_000;

    // ====== Исключаемые разделы ======
    private static final Set<String> EXCLUDED_SECTIONS = Set.of(
            "Подарочные сертификаты",
            "Автохимия, детейлинг",
            "Все для стеклопластика",
            "Кабель и комплектующие для автоакустики",
            "Ремкомплекты",
            "Шумоизоляция",
            "Виброизоляция",
            "Раздел не для выгрузки",
            "Фильтра Базальт",
            "Сувенирная продукция"
    );

    // Кеши для pathName и групп контрагентов
    private static final Map<String, String> PRODUCT_PATH_CACHE = new HashMap<>();
    private static final Map<String, String> FOLDER_PATH_CACHE = new HashMap<>();
    private static final Map<String, String> agentGroupCache = new HashMap<>();
    private static final Map<String, Boolean> RETAIL_BUYER_CACHE = new HashMap<>();

    // Формат денег
    private static final DecimalFormat MONEY_FMT;
    static {
        DecimalFormatSymbols s = new DecimalFormatSymbols(new Locale("ru", "RU"));
        s.setGroupingSeparator(' ');
        MONEY_FMT = new DecimalFormat("#,##0.00 ₽", s);
    }

    // внутренняя структура для неликвидов
    private static class StaleItem {
        String href;
        String id;      // <-- ВАЖНО: UUID (ключ для продаж)
        String name;
        String code;
        String pathName;

        double stock;
        double price;
        double cost;

        long daysOnStock;

        double soldQty;
        double dailySales;
        double coverageDays;

        String category; // "МЕРТВЕЦ", "ОЧЕНЬ МЕДЛЕННЫЙ", "" ...
    }

    public static void main(String[] args) throws Exception {
        System.out.println("Retail group href = " + RETAIL_GROUP_HREF);

        LocalDate today = LocalDate.now();
        YearMonth ym = YearMonth.from(today);
        LocalDate monthStart = ym.atDay(1);
        LocalDate monthEnd = ym.atEndOfMonth();

        // Отчетные продажи: последние 180 дней до конца предыдущего месяца
        LocalDate salesTo = today.withDayOfMonth(1).minusDays(1);
        LocalDate salesFrom = salesTo.minusDays(SALES_LOOKBACK_DAYS - 1);

        System.out.println("Сегодня: " + today);
        System.out.println("Отчетные продажи: " + salesFrom + " — " + salesTo);
        System.out.println("Исключаем pathName: " + EXCLUDED_SECTIONS);
        System.out.println();

        String storeName = "Основной склад";
        String storeHref = findStoreHrefByName(storeName);
        if (storeHref == null) {
            System.out.println("Склад \"" + storeName + "\" не найден. Завершаем.");
            return;
        }
        System.out.println("Найден склад: " + storeName);
        System.out.println("storeHref: " + storeHref);
        System.out.println("--------------------------------------------");

        // === 1. Продажи за 180 дней: demand + retaildemand (все покупатели), ключ = UUID ===
        Map<String, Double> salesQtyById = loadSalesMapById(salesFrom, salesTo);

        System.out.println("Продаж собрано (по UUID ассортимента): " + salesQtyById.size());
        System.out.println();

        // === 2. Остатки/неликвид по одному складу ===
        List<StaleItem> staleItems = new ArrayList<>();
        double totalStockCost = 0.0;
        double staleStockCost = 0.0;

        int page = 0;
        int pageSize = 100;

        while (true) {
            page++;
            int offset = (page - 1) * pageSize;

            String stockUrl = BASE_URL + "/report/stock/all"
                    + "?limit=" + pageSize
                    + "&offset=" + offset
                    + "&store=" + urlEncode(storeHref);

            String json = get(stockUrl);

            JsonNode root;
            try {
                root = MAPPER.readTree(json);
            } catch (JsonProcessingException e) {
                System.out.println("stock/all: неожиданный ответ, не JSON. Текст:\n" + json);
                break;
            }

            JsonNode rows = root.path("rows");
            if (!rows.isArray() || rows.size() == 0) {
                if (page == 1) {
                    System.out.println("stock/all: неожиданный ответ, нет массива rows. JSON:");
                    System.out.println(json);
                }
                break;
            }

            System.out.println("Страница stock/all: " + page);
            int addedStale = 0;

            for (JsonNode row : rows) {

                JsonNode meta = row.path("meta");
                String type = meta.path("type").asText("");

                if (!"product".equals(type) && !"variant".equals(type)) {
                    continue;
                }

                String href = meta.path("href").asText("");
                if (href.isEmpty()) continue;

                String id = extractId(href);
                if (id.isBlank()) continue;

                String name = row.path("name").asText("");
                String code = row.path("code").asText("");

                double stock = row.path("stock").asDouble(0);
                if (stock <= 0) continue;

                double salePriceRaw = row.path("salePrice").asDouble(0);
                double priceRaw = row.path("price").asDouble(0);
                double raw = salePriceRaw > 0 ? salePriceRaw : priceRaw;
                double price = raw / 100.0;
                if (price <= 0) continue;

                long daysOnStock = row.path("stockDays").asLong(0);

                double lineCost = stock * price;
                totalStockCost += lineCost;

                if (daysOnStock < MIN_AGE_DAYS) continue;

                String path = resolveProductPath(href);
                if (isExcludedBySection(path)) continue;

                // ВАЖНО: продажи берём по UUID (а не по href) — это чинит "везде 0"
                double soldQty = salesQtyById.getOrDefault(id, 0.0);
                double dailySales = soldQty / SALES_LOOKBACK_DAYS;
                double coverageDays = (dailySales > 0) ? stock / dailySales : Double.POSITIVE_INFINITY;

                if (dailySales > 0 && coverageDays < FAST_TURNOVER_DAYS) {
                    continue;
                }

                StaleItem item = new StaleItem();
                item.href = href;
                item.id = id;
                item.name = name;
                item.code = code;
                item.pathName = path;
                item.stock = stock;
                item.price = price;
                item.cost = lineCost;
                item.daysOnStock = daysOnStock;
                item.soldQty = soldQty;
                item.dailySales = dailySales;
                item.coverageDays = coverageDays;
                item.category = classify(item);

                staleItems.add(item);
                staleStockCost += lineCost;
                addedStale++;
            }

            System.out.println("Добавлено в неликвид: " + addedStale);
            System.out.println();

            if (rows.size() < pageSize) break;
        }

        if (staleItems.isEmpty()) {
            System.out.println("Неликвидов нет, отчёты делать не из чего.");
            return;
        }

        staleItems.sort(Comparator
                .comparingLong((StaleItem it) -> it.daysOnStock).reversed()
                .thenComparingDouble(it -> it.dailySales));

        double staleShare = totalStockCost > 0 ? (staleStockCost / totalStockCost) * 100.0 : 0.0;

        System.out.println("============ ИТОГО ПО СКЛАДУ ============");
        System.out.printf("Общая стоимость склада: %s%n", MONEY_FMT.format(totalStockCost));
        System.out.printf("Стоимость неликвида:   %s%n", MONEY_FMT.format(staleStockCost));
        System.out.printf("Доля неликвида:        %.2f%%%n", staleShare);
        System.out.println("==========================================");
        System.out.println();

        System.out.println("============ ОБОРОТ МАГАЗИНА (розница) ============");
        System.out.printf("Оборот за %d дней (retaildemand, все склады): %s%n",
                SALES_LOOKBACK_DAYS, MONEY_FMT.format(totalTurnover180));

        double avgMonthlyTurnover = totalTurnover180 / MONTHS_IN_PERIOD;
        System.out.printf("Среднемесячный оборот (розница): %s%n", MONEY_FMT.format(avgMonthlyTurnover));

        double planShare = pickPlanShare(staleShare);
        System.out.printf("Рекомендованная доля плана по неликвиду: %.1f%%%n", planShare * 100.0);
        System.out.println(" (при низкой доле неликвида ближе к 5%, при высокой — до 35%)");
        System.out.println("==========================================");
        System.out.println();

        double monthlyPlanByHorizon = staleStockCost / 12.0;

        double monthlyPlanByTurnover;
        boolean turnoverBased;

        if (avgMonthlyTurnover > 0) {
            monthlyPlanByTurnover = avgMonthlyTurnover * planShare;
            turnoverBased = (monthlyPlanByTurnover <= monthlyPlanByHorizon);
        } else {
            monthlyPlanByTurnover = monthlyPlanByHorizon;
            turnoverBased = false;
        }

        double monthlyPlan = Math.min(monthlyPlanByHorizon, monthlyPlanByTurnover);

        System.out.println("------ ПЛАН ПО НЕЛИКВИДУ НА МЕСЯЦ ------");
        if (turnoverBased) {
            System.out.printf("Итоговый план ограничен оборотом (%.1f%% от среднемесячного оборота).%n", planShare * 100.0);
        } else {
            System.out.println("Итоговый план ограничен горизонтом (равномерное распределение неликвида на 12 месяцев).");
        }
        System.out.printf("Итоговый план продаж неликвида на месяц: %s%n", MONEY_FMT.format(monthlyPlan));

        double horizonMonths = (monthlyPlan <= 0.0) ? Double.POSITIVE_INFINITY : (staleStockCost / monthlyPlan);
        if (Double.isInfinite(horizonMonths)) {
            System.out.println("При таком плане срок очистки склада оценить нельзя (план = 0).");
        } else {
            System.out.printf("Срок полной очистки неликвида при таком плане: ≈ %.1f месяцев%n", horizonMonths);
        }
        System.out.println("-----------------------------------------");
        System.out.println();

        // === Список для ПЛАНА на месяц (под monthlyPlan) ===
        List<StaleItem> planItems = new ArrayList<>();
        double planSum = 0.0;
        for (StaleItem it : staleItems) {
            if (!"МЕРТВЕЦ".equals(it.category) && !"ОЧЕНЬ МЕДЛЕННЫЙ".equals(it.category)) continue;
            if (planSum >= monthlyPlan) break;
            if (it.cost <= 0) continue;
            planItems.add(it);
            planSum += it.cost;
        }

        System.out.println("План на месяц (по неликвиду):");
        System.out.println("   План по деньгам: " + MONEY_FMT.format(monthlyPlan));
        System.out.println("   Набрано в план (по позициям): " + MONEY_FMT.format(planSum));
        System.out.println("   Позиции в плане: " + planItems.size());
        System.out.println();

        // === Подбор для розницы (по бюджету) — печать в консоль ===
        double budgetLeft = RETAIL_BUDGET;
        List<StaleItem> picked = new ArrayList<>();

        for (StaleItem it : staleItems) {
            if (budgetLeft <= 0) break;
            if (it.cost <= 0) continue;
            if (it.cost > budgetLeft) continue;

            picked.add(it);
            budgetLeft -= it.cost;
        }

        double pickedSum = picked.stream().mapToDouble(it -> it.cost).sum();

        System.out.println("====== ПОДБОР ДЛЯ РОЗНИЦЫ (по бюджету) ======");
        System.out.println("Бюджет: " + MONEY_FMT.format(RETAIL_BUDGET));
        System.out.println("Выбрано позиций: " + picked.size());
        System.out.println("Сумма: " + MONEY_FMT.format(pickedSum));
        System.out.println();

        int idx = 1;
        for (StaleItem it : picked) {
            System.out.printf("%d. %s (код %s), %.1f шт%n", idx++, it.name, it.code, it.stock);
            System.out.printf("   %d дней на складе%n", it.daysOnStock);
            System.out.printf("   Цена: %.2f ₽ | Сумма: %.2f ₽%n", it.price, it.cost);
            System.out.printf("   Продано: %.1f шт%n", it.soldQty);
            System.out.printf("   Средние/день: %.3f%n", it.dailySales);
            System.out.printf("   Покрытие: %s%n", Double.isInfinite(it.coverageDays) ? "∞" : String.format(Locale.US, "%.0f", it.coverageDays));
            System.out.println();
        }

        // === Excel: Файл 1 — план на месяц (для менеджеров) ===
        String planFileName = resolveFreeFileName("stale_plan_" + ym + ".xlsx");
        try {
            writePlanExcel(
                    planFileName,
                    monthlyPlan,
                    planItems,
                    planSum,
                    monthStart,
                    monthEnd,
                    salesFrom,
                    salesTo
            );
            System.out.println("Excel-файл плана на месяц записан: " + planFileName);
        } catch (IOException e) {
            System.out.println("Ошибка записи Excel (план на месяц): " + e.getMessage());
        }

        // === Excel: Файл 2 — полный неликвид ===
        String fullFileName = resolveFreeFileName("stale_full_" + ym + ".xlsx");
        try {
            writeFullExcel(
                    fullFileName,
                    staleItems,
                    staleStockCost,
                    monthStart,
                    monthEnd,
                    salesFrom,
                    salesTo
            );
            System.out.println("Excel-файл полного неликвида записан: " + fullFileName);
        } catch (IOException e) {
            System.out.println("Ошибка записи Excel (полный неликвид): " + e.getMessage());
        }

        // === Отправляем оба файла на почту ===
        try {
            sendEmailWithAttachments(
                    "Отчёты по неликвиду за " + today,
                    "Во вложении:\n" +
                            "1) План продаж неликвида на месяц\n" +
                            "2) Полный список неликвида\n\n" +
                            "В этом списке – товары, которые нужно продать в этом месяце.\n" +
                            "Премия = 10% от розничной цены, но не более 4 000 ₽ за единицу.\n",
                    Arrays.asList(planFileName, fullFileName)
            );
            System.out.println("Письмо с отчётами отправлено на " + EMAIL_TO);
        } catch (Exception e) {
            System.out.println("⚠ Ошибка отправки e-mail: " + e.getMessage());
            e.printStackTrace(System.out);
        }
    }

    // =====================================================================
    // Классификация
    // =====================================================================
    private static String classify(StaleItem it) {
        if (it.soldQty <= 0.000001) return "МЕРТВЕЦ";
        if (Double.isInfinite(it.coverageDays)) return "МЕРТВЕЦ";
        if (it.coverageDays >= 180.0) return "ОЧЕНЬ МЕДЛЕННЫЙ";
        return "";
    }

    private static double pickPlanShare(double staleSharePercent) {
        if (staleSharePercent <= 0.0) return 0.0;
        double capped = Math.min(staleSharePercent, 50.0);
        double k = capped / 50.0;
        return MIN_STALE_PLAN_SHARE + k * (MAX_STALE_PLAN_SHARE - MIN_STALE_PLAN_SHARE);
    }

    // =====================================================================
    // Ключевая правка: продажи считаем по UUID, а не по href
    // =====================================================================
    private static String extractId(String href) {
        if (href == null || href.isBlank()) return "";
        String h = href;
        int q = h.indexOf('?');
        if (q >= 0) h = h.substring(0, q);
        int slash = h.lastIndexOf('/');
        if (slash < 0 || slash == h.length() - 1) return "";
        return h.substring(slash + 1).trim();
    }

    // =====================================================================
    // Поиск склада по имени
    // =====================================================================
    private static String findStoreHrefByName(String storeName)
            throws IOException, InterruptedException, URISyntaxException {

        String url = BASE_URL + "/entity/store?limit=1000";
        String json = get(url);

        JsonNode root = MAPPER.readTree(json);
        JsonNode rows = root.path("rows");
        if (!rows.isArray()) {
            System.out.println("store: неожиданный ответ, нет rows");
            return null;
        }

        for (JsonNode row : rows) {
            String name = row.path("name").asText("");
            if (storeName.equals(name)) {
                return row.path("meta").path("href").asText("");
            }
        }
        return null;
    }

    // =====================================================================
    // Загрузка продаж: demand + retaildemand, ключ = UUID
    // =====================================================================
    private static Map<String, Double> loadSalesMapById(LocalDate from, LocalDate to)
            throws IOException, InterruptedException, URISyntaxException {

        Map<String, Double> result = new HashMap<>();

        String dateFrom = from.toString() + " 00:00:00";
        String dateTo = to.toString() + " 23:59:59";

        loadSalesFromEntityById("demand", dateFrom, dateTo, result);
        loadSalesFromEntityById("retaildemand", dateFrom, dateTo, result);

        return result;
    }

    private static void loadSalesFromEntityById(String entityName,
                                                String dateFrom,
                                                String dateTo,
                                                Map<String, Double> salesMapById)
            throws IOException, InterruptedException, URISyntaxException {

        System.out.println();
        System.out.println("=== Загрузка продаж из /entity/" + entityName + " ===");

        int limit = 100;
        int offset = 0;
        int page = 0;
        int totalDocs = 0;

        while (true) {
            page++;

            String rawFilter = "moment>=" + dateFrom +
                    ";moment<=" + dateTo +
                    ";applicable=true";

            String filter = URLEncoder.encode(rawFilter, StandardCharsets.UTF_8);

            String url = BASE_URL + "/entity/" + entityName
                    + "?limit=" + limit
                    + "&offset=" + offset
                    + "&filter=" + filter;

            String json = get(url);

            JsonNode root;
            try {
                root = MAPPER.readTree(json);
            } catch (JsonProcessingException e) {
                System.out.println(entityName + ": неожиданный ответ, не JSON. Текст:\n" + json);
                break;
            }

            JsonNode rows = root.path("rows");
            if (!rows.isArray() || rows.size() == 0) {
                break;
            }

            System.out.println("  Страница " + page + ": " + rows.size());
            totalDocs += rows.size();

            for (JsonNode doc : rows) {
                String docHref = doc.path("meta").path("href").asText("");
                if (docHref.isEmpty()) continue;

                boolean countAsRetailTurnover = "retaildemand".equals(entityName);

                loadPositionsForDocById(docHref, entityName, salesMapById, countAsRetailTurnover);
            }

            if (rows.size() < limit) break;
            offset += limit;
        }

        System.out.println("Всего документов: " + totalDocs);
    }

    private static void loadPositionsForDocById(String docHref,
                                                String entityName,
                                                Map<String, Double> salesMapById,
                                                boolean countAsRetailTurnover)
            throws IOException, InterruptedException, URISyntaxException {
        if ("retaildemand".equals(entityName)) {
            System.out.println("retaildemand doc=" + docHref + " countAsRetailTurnover=" + countAsRetailTurnover);
        }
        int limit = 100;
        int offset = 0;

        while (true) {
            String url = docHref + "/positions?limit=" + limit + "&offset=" + offset;

            String json = get(url);

            JsonNode root;
            try {
                root = MAPPER.readTree(json);
            } catch (JsonProcessingException e) {
                System.out.println("  " + entityName + " positions: неожиданный ответ, не JSON. Текст:\n" + json);
                break;
            }

            JsonNode rows = root.path("rows");
            if (!rows.isArray() || rows.size() == 0) {
                break;
            }

            for (JsonNode position : rows) {
                String href = position.path("assortment").path("meta").path("href").asText("");
                if (href.isBlank()) continue;

                String id = extractId(href);
                if (id.isBlank()) continue;

                double quantity = position.path("quantity").asDouble(0.0);

                // продажи всегда по всем покупателям
                salesMapById.merge(id, quantity, Double::sum);

                // оборот только для retaildemand + розничная группа
                if ("retaildemand".equals(entityName) && countAsRetailTurnover) {
                    double lineSumRaw = position.path("sum").asDouble(0.0);
                    double priceRaw = position.path("price").asDouble(0.0);

                    double lineTurnover = (lineSumRaw > 0)
                            ? (lineSumRaw / 100.0)
                            : (priceRaw / 100.0) * quantity;

                    totalTurnover180 += lineTurnover;
                }
            }

            if (rows.size() < limit) break;
            offset += limit;
        }
    }

    // =====================================================================
    // Розничный покупатель по ГРУППЕ (agentGroup.name)
    // =====================================================================

    private static boolean isRetailCustomer(String agentHref)
            throws IOException, InterruptedException, URISyntaxException {

        if (agentHref == null || agentHref.isBlank()) return false;

        Boolean cached = RETAIL_BUYER_CACHE.get(agentHref);
        if (cached != null) return cached;

        String json = get(agentHref);
        JsonNode root = MAPPER.readTree(json);

        boolean isRetail = false;

        // 1) Основной критерий: группа контрагента
        String groupHref = root.path("group").path("meta").path("href").asText("");
        if (!groupHref.isBlank() && groupHref.equals(RETAIL_GROUP_HREF)) {
            isRetail = true;
        }

        // 2) Подстраховка: имя контрагента (как в UI)
        if (!isRetail) {
            String agentName = root.path("name").asText("");
            if (RETAIL_GROUP_NAME.equalsIgnoreCase(agentName)) {
                isRetail = true;
            }
        }

        RETAIL_BUYER_CACHE.put(agentHref, isRetail);
        return isRetail;
    }

    private static String getAgentGroupName(String agentHref)
            throws IOException, InterruptedException, URISyntaxException {

        if (agentGroupCache.containsKey(agentHref)) {
            return agentGroupCache.get(agentHref);
        }

        String agentJson = get(agentHref);
        JsonNode agent = MAPPER.readTree(agentJson);
        String groupHref = agent.path("agentGroup").path("meta").path("href").asText("");

        String groupName = "";
        if (!groupHref.isEmpty()) {
            String groupJson = get(groupHref);
            JsonNode group = MAPPER.readTree(groupJson);
            groupName = group.path("name").asText("");
        }

        agentGroupCache.put(agentHref, groupName);
        return groupName;
    }

    // =====================================================================
    // PATH RESOLVER
    // =====================================================================
    private static String resolveProductPath(String productHref) {
        if (productHref == null || productHref.isBlank()) return "";

        if (PRODUCT_PATH_CACHE.containsKey(productHref)) {
            return PRODUCT_PATH_CACHE.get(productHref);
        }

        String result = "";

        try {
            String url = productHref + "?fields=pathName,productFolder";
            String json = get(url);
            JsonNode node = MAPPER.readTree(json);

            String direct = node.path("pathName").asText("");
            if (!direct.isBlank()) {
                result = direct;
            } else {
                String folderHref = node.path("productFolder").path("meta").path("href").asText("");
                if (!folderHref.isBlank()) {
                    result = resolveFolderPath(folderHref);
                }
            }
        } catch (Exception e) {
            System.err.println("Ошибка resolveProductPath: " + e.getMessage());
        }

        PRODUCT_PATH_CACHE.put(productHref, result);
        return result;
    }

    private static String resolveFolderPath(String folderHref)
            throws IOException, URISyntaxException, InterruptedException {

        if (folderHref == null || folderHref.isBlank()) return "";

        if (FOLDER_PATH_CACHE.containsKey(folderHref)) {
            return FOLDER_PATH_CACHE.get(folderHref);
        }

        String json = get(folderHref);
        JsonNode node = MAPPER.readTree(json);

        String name = node.path("name").asText("");
        String parentPath = node.path("pathName").asText("");

        String result = !parentPath.isBlank() ? parentPath + "/" + name : name;

        FOLDER_PATH_CACHE.put(folderHref, result);
        return result;
    }

    private static boolean isExcludedBySection(String fullPath) {
        if (fullPath == null || fullPath.isBlank()) return false;
        for (String ex : EXCLUDED_SECTIONS) {
            if (fullPath.contains(ex)) return true;
        }
        return false;
    }

    // =====================================================================
    // Excel: план
    // =====================================================================

// 3) ЗАМЕНИ СИГНАТУРУ writePlanExcel(...) и добавь строки в "шапку"

    private static void writePlanExcel(
            String fileName,
            double monthlyPlan,
            List<StaleItem> planItems,
            double planSum,
            LocalDate planFrom,
            LocalDate planTo,
            LocalDate salesFrom,
            LocalDate salesTo
    ) throws IOException {

        try (Workbook wb = new XSSFWorkbook()) {
            DataFormat df = wb.createDataFormat();

            CellStyle moneyStyle = wb.createCellStyle();
            moneyStyle.setDataFormat(df.getFormat("#,##0.00"));

            CellStyle intStyle = wb.createCellStyle();
            intStyle.setDataFormat(df.getFormat("#,##0"));

            Sheet sheet = wb.createSheet("План неликвида");

            int r = 0;

            Row row0 = sheet.createRow(r++);
            row0.createCell(0).setCellValue("Период плана:");
            row0.createCell(1).setCellValue(planFrom + " — " + planTo);

            Row row0b = sheet.createRow(r++);
            row0b.createCell(0).setCellValue("Период продаж (180 дней):");
            row0b.createCell(1).setCellValue(salesFrom + " — " + salesTo);

            r++;

            Row row1 = sheet.createRow(r++);
            row1.createCell(0).setCellValue("План продаж неликвида на месяц");
            Cell planCell = row1.createCell(1);
            planCell.setCellValue(monthlyPlan);
            planCell.setCellStyle(moneyStyle);

            Row row2 = sheet.createRow(r++);
            row2.createCell(0).setCellValue("Набрано в план (по позициям), ₽");
            Cell planSumCell = row2.createCell(1);
            planSumCell.setCellValue(planSum);
            planSumCell.setCellStyle(moneyStyle);

            double coverage = (monthlyPlan > 0) ? (planSum / monthlyPlan * 100.0) : 0.0;
            Row row3 = sheet.createRow(r++);
            row3.createCell(0).setCellValue("Покрытие плана, %");
            row3.createCell(1).setCellValue(coverage);

            r++;

            Row row4 = sheet.createRow(r++);
            row4.createCell(0).setCellValue(
                    "В этом списке – товары, которые нужно продать в этом месяце. " +
                            "Премия = 10% от розничной цены, но не более 4 000 ₽ за единицу."
            );

            r++;

            Row header = sheet.createRow(r++);
            header.createCell(0).setCellValue("Наименование");
            header.createCell(1).setCellValue("Категория");
            header.createCell(2).setCellValue("Дней на складе");
            header.createCell(3).setCellValue("Остаток, шт");
            header.createCell(4).setCellValue("Продано за 180 дней, шт");
            header.createCell(5).setCellValue("Хватит на, дней");
            header.createCell(6).setCellValue("Цена, ₽");

            for (StaleItem it : planItems) {
                Row row = sheet.createRow(r++);

                row.createCell(0).setCellValue(it.name);
                row.createCell(1).setCellValue(it.category);

                Cell cDays = row.createCell(2);
                cDays.setCellValue(it.daysOnStock);
                cDays.setCellStyle(intStyle);

                row.createCell(3).setCellValue(it.stock);
                row.createCell(4).setCellValue(it.soldQty);

                Cell cCover = row.createCell(5);
                if (Double.isInfinite(it.coverageDays)) {
                    cCover.setCellValue("∞");
                } else {
                    cCover.setCellValue(Math.round(it.coverageDays));
                    cCover.setCellStyle(intStyle);
                }

                Cell cPrice = row.createCell(6);
                cPrice.setCellValue(it.price);
                cPrice.setCellStyle(moneyStyle);
            }

            for (int col = 0; col <= 6; col++) sheet.autoSizeColumn(col);

            try (FileOutputStream fos = new FileOutputStream(fileName)) {
                wb.write(fos);
            }
        }
    }

    // =====================================================================
    // Excel: полный список
    // =====================================================================
    // 4) ЗАМЕНИ СИГНАТУРУ writeFullExcel(...) и добавь строки в "шапку"

    private static void writeFullExcel(
            String fileName,
            List<StaleItem> staleItems,
            double totalStaleCost,
            LocalDate planFrom,
            LocalDate planTo,
            LocalDate salesFrom,
            LocalDate salesTo
    ) throws IOException {

        try (Workbook wb = new XSSFWorkbook()) {
            DataFormat df = wb.createDataFormat();

            CellStyle moneyStyle = wb.createCellStyle();
            moneyStyle.setDataFormat(df.getFormat("#,##0.00"));

            CellStyle intStyle = wb.createCellStyle();
            intStyle.setDataFormat(df.getFormat("#,##0"));

            Sheet sheet = wb.createSheet("Полный неликвид");

            int r = 0;

            Row title = sheet.createRow(r++);
            title.createCell(0).setCellValue("ПОЛНЫЙ СПИСОК НЕЛИКВИДНОГО ТОВАРА");

            Row row0 = sheet.createRow(r++);
            row0.createCell(0).setCellValue("Период плана:");
            row0.createCell(1).setCellValue(planFrom + " — " + planTo);

            Row row0b = sheet.createRow(r++);
            row0b.createCell(0).setCellValue("Период продаж (180 дней):");
            row0b.createCell(1).setCellValue(salesFrom + " — " + salesTo);

            Row info = sheet.createRow(r++);
            info.createCell(0).setCellValue(
                    "«Хватит на, дней» — это сколько дней хватит текущего остатка при той скорости продаж, " +
                            "которая была за последние 180 дней. Если указано \"∞\" — товар за 180 дней ни разу не продавался."
            );

            r++;

            Row header = sheet.createRow(r++);
            header.createCell(0).setCellValue("Наименование");
            header.createCell(1).setCellValue("Категория");
            header.createCell(2).setCellValue("Дней на складе");
            header.createCell(3).setCellValue("Остаток, шт");
            header.createCell(4).setCellValue("Цена, ₽");
            header.createCell(5).setCellValue("Сумма, ₽");
            header.createCell(6).setCellValue("Продано за 180 дней, шт");
            header.createCell(7).setCellValue("Хватит на, дней");
            header.createCell(8).setCellValue("Категория (pathName)");

            for (StaleItem it : staleItems) {
                Row row = sheet.createRow(r++);

                row.createCell(0).setCellValue(it.name);
                row.createCell(1).setCellValue(it.category);

                Cell cDays = row.createCell(2);
                cDays.setCellValue(it.daysOnStock);
                cDays.setCellStyle(intStyle);

                row.createCell(3).setCellValue(it.stock);

                Cell cPrice = row.createCell(4);
                cPrice.setCellValue(it.price);
                cPrice.setCellStyle(moneyStyle);

                Cell cCost = row.createCell(5);
                cCost.setCellValue(it.cost);
                cCost.setCellStyle(moneyStyle);

                row.createCell(6).setCellValue(it.soldQty);

                Cell cCover = row.createCell(7);
                if (Double.isInfinite(it.coverageDays)) {
                    cCover.setCellValue("∞");
                } else {
                    cCover.setCellValue(Math.round(it.coverageDays));
                    cCover.setCellStyle(intStyle);
                }

                row.createCell(8).setCellValue(it.pathName);
            }

            Row totalRow = sheet.createRow(r++);
            totalRow.createCell(0).setCellValue("ИТОГО стоимости неликвида, ₽");
            Cell cTotal = totalRow.createCell(5);
            cTotal.setCellValue(totalStaleCost);
            cTotal.setCellStyle(moneyStyle);

            Row countRow = sheet.createRow(r++);
            countRow.createCell(0).setCellValue("Количество позиций неликвида");
            countRow.createCell(1).setCellValue(staleItems.size());

            for (int col = 0; col <= 8; col++) sheet.autoSizeColumn(col);

            try (FileOutputStream fos = new FileOutputStream(fileName)) {
                wb.write(fos);
            }
        }
    }


    // =====================================================================
    // Свободное имя файла
    // =====================================================================
    private static String resolveFreeFileName(String baseName) {
        File f = new File(baseName);
        if (!f.exists()) return baseName;

        String prefix;
        String suffix;
        int dot = baseName.lastIndexOf('.');
        if (dot == -1) {
            prefix = baseName;
            suffix = "";
        } else {
            prefix = baseName.substring(0, dot);
            suffix = baseName.substring(dot);
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

    // =====================================================================
    // E-mail
    // =====================================================================
    private static void sendEmailWithAttachments(String subject,
                                                 String bodyText,
                                                 List<String> fileNames) throws Exception {

        Properties props = new Properties();
        props.put("mail.smtp.auth", "true");
        props.put("mail.smtp.ssl.enable", "true");
        props.put("mail.smtp.host", SMTP_HOST);
        props.put("mail.smtp.port", String.valueOf(SMTP_PORT));
        props.put("mail.debug", "false");

        Session session = Session.getInstance(props, new jakarta.mail.Authenticator() {
            @Override
            protected PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication(SMTP_USER, SMTP_PASSWORD);
            }
        });

        Message message = new MimeMessage(session);
        message.setFrom(new InternetAddress(EMAIL_FROM));
        message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(EMAIL_TO));
        message.setSubject(subject);

        MimeBodyPart textPart = new MimeBodyPart();
        textPart.setText(bodyText, "UTF-8");

        Multipart multipart = new MimeMultipart();
        multipart.addBodyPart(textPart);

        for (String fn : fileNames) {
            MimeBodyPart attachPart = new MimeBodyPart();
            attachPart.attachFile(new File(fn));
            multipart.addBodyPart(attachPart);
        }

        message.setContent(multipart);

        Transport.send(message);
        System.out.println("✔ Письмо отправлено через Mail.ru");
    }

    // =====================================================================
    // HTTP GET
    // =====================================================================
    private static String get(String url) throws IOException, InterruptedException, URISyntaxException {
        HttpRequest request = HttpRequest.newBuilder()
                .uri(new URI(url))
                .header("Authorization", basicAuthHeader())
                .header("Accept", "application/json;charset=utf-8")
                .header("Accept-Encoding", "gzip")
                .GET()
                .build();

        int max429Attempts = 5;
        int attempt429 = 0;
        HttpResponse<byte[]> response;

        while (true) {
            response = sendWithRetry(request, url);
            int status = response.statusCode();

            if (status == 429) {
                attempt429++;

                String body429 = unpack(response);
                System.out.println("HTTP 429 (rate limit) при запросе: " + url);
                System.out.println("Тело ответа:");
                System.out.println(body429);
                System.out.println("Попытка " + attempt429 + " из " + max429Attempts);

                if (attempt429 >= max429Attempts) {
                    System.out.println("Превышено число попыток при 429, продолжаем с этим ответом.");
                    break;
                }

                long sleepMs = 3000L;
                Optional<String> retryAfter = response.headers().firstValue("Retry-After");
                if (retryAfter.isPresent()) {
                    try {
                        sleepMs = Long.parseLong(retryAfter.get()) * 1000L;
                    } catch (NumberFormatException ignored) {}
                }

                Thread.sleep(sleepMs);
                continue;
            }

            break;
        }

        String body = unpack(response);
        int status = response.statusCode();
        if (status != 200) {
            System.out.println("HTTP " + status + " при запросе: " + url);
            System.out.println("Тело ответа:");
            System.out.println(body);
        }

        Thread.sleep(50L);
        return body;
    }

    private static HttpResponse<byte[]> sendWithRetry(HttpRequest request, String desc)
            throws IOException, InterruptedException {
        int maxAttempts = 3;
        for (int attempt = 1; attempt <= maxAttempts; attempt++) {
            try {
                return httpClient.send(request, HttpResponse.BodyHandlers.ofByteArray());
            } catch (IOException e) {
                System.out.println("   ⚠ Ошибка соединения (" + desc + "), попытка "
                        + attempt + "/" + maxAttempts + ": " + e.getMessage());
                if (attempt == maxAttempts) throw e;
                Thread.sleep(2000L * attempt);
            }
        }
        throw new IOException("Не удалось выполнить запрос: " + desc);
    }

    private static String unpack(HttpResponse<byte[]> response) throws IOException {
        byte[] bytes = response.body();

        Optional<String> enc = response.headers().firstValue("Content-Encoding");
        if (enc.isPresent() && enc.get().contains("gzip")) {
            try (GZIPInputStream gis = new GZIPInputStream(new ByteArrayInputStream(bytes));
                 InputStreamReader isr = new InputStreamReader(gis, StandardCharsets.UTF_8);
                 BufferedReader br = new BufferedReader(isr)) {

                StringBuilder sb = new StringBuilder();
                String line;
                while ((line = br.readLine()) != null) sb.append(line);
                return sb.toString();
            }
        }
        return new String(bytes, StandardCharsets.UTF_8);
    }

    private static String basicAuthHeader() {
        String auth = LOGIN + ":" + PASSWORD;
        String encoded = Base64.getEncoder().encodeToString(auth.getBytes(StandardCharsets.UTF_8));
        return "Basic " + encoded;
    }

    private static String urlEncode(String s) {
        return URLEncoder.encode(s, StandardCharsets.UTF_8);
    }
}
