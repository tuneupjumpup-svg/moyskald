package org.example;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import java.io.BufferedReader;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.net.URI;
import java.net.URISyntaxException;
import java.net.URLEncoder;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.charset.StandardCharsets;
import java.time.LocalDate;
import java.util.Base64;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.util.zip.GZIPInputStream;

public class CheckProductRetailSales {

    // ==== Авторизация MoySklad (из рабочего кода) ====
    private static final String LOGIN = "smirnova@timokhovta5";
    private static final String PASSWORD = "cascada3788932";
    private static final String BASE_URL = "https://api.moysklad.ru/api/remap/1.2";
    // Период проверки (кол-во дней до даты отгрузки)

    // Дата начала проверки
    private static final LocalDate FROM_DATE = LocalDate.of(2025, 9, 7);

    // Дата отгрузки (конец периода)
    private static final LocalDate TO_DATE = LocalDate.of(2025, 9, 8);

    // Искомый товар
    private static final String TARGET_PRODUCT_NAME = "СЧ Динамики AZ-13 SPL Power DJ 6.5 V2";

    // Дата отгрузки, от которой считаем 180 дней назад
    private static final LocalDate SHIPMENT_DATE = LocalDate.of(2025, 9, 8);



    private static final ObjectMapper MAPPER = new ObjectMapper();
    private static final HttpClient httpClient = HttpClient.newBuilder()
            .followRedirects(HttpClient.Redirect.ALWAYS)
            .build();

    // Кэш: href контрагента → является ли розничным покупателем
    private static final Map<String, Boolean> RETAIL_BUYER_CACHE = new HashMap<>();

    // Кэш: href товара → его name
    private static final Map<String, String> PRODUCT_NAME_CACHE = new HashMap<>();

    // Список заказов, где найден товар: API href → qty
    private static final Map<String, Double> RETAIL_ORDERS = new LinkedHashMap<>();

    public static void main(String[] args) throws Exception {
        LocalDate from = FROM_DATE;
        LocalDate to = TO_DATE;

        System.out.println("Дата отгрузки товара: " + TO_DATE);
        System.out.println("Проверяем продажи за период: " + from + " — " + to);
        System.out.println("Товар: " + TARGET_PRODUCT_NAME);
        System.out.println("--------------------------------------------");

        String dateFrom = from.toString() + " 00:00:00";
        String dateTo = to.toString() + " 23:59:59";

        // считаем продажи по demand + retaildemand, по всем покупателям
        double soldQty = calcSalesForProduct(TARGET_PRODUCT_NAME, dateFrom, dateTo);

        System.out.println();
        if (soldQty > 0) {
            System.out.printf(
                    "РЕЗУЛЬТАТ: товар ПРОДАВАЛСЯ за период %s — %s.%n",
                    from, to
            );
            System.out.printf(Locale.US, "Всего продано: %.2f шт%n", soldQty);
        } else {
            System.out.printf(
                    "РЕЗУЛЬТАТ: товар НЕ ПРОДАВАЛСЯ за период %s — %s.%n",
                    from, to
            );
        }

        if (!RETAIL_ORDERS.isEmpty()) {
            System.out.println();
            System.out.println("Документы, где был продан товар (веб-ссылки МойСклад):");
            RETAIL_ORDERS.forEach((apiHref, qty) -> {
                String webLink = buildWebLinkFromApiHref(apiHref);
                if (webLink != null && !webLink.isBlank()) {
                    System.out.printf(" %s   qty=%.2f%n", webLink, qty);
                }
            });
        }
    }



    /**
     * Основной метод: возвращает суммарное количество, проданное в розницу за период.
     */
    // Подсчёт продаж по двум сущностям: demand + retaildemand
    private static double calcSalesForProduct(String productName,
                                              String dateFrom,
                                              String dateTo)
            throws IOException, InterruptedException, URISyntaxException {

        double total = 0.0;

        // 1) Отгрузка (demand)
        total += calcSalesFromEntity("demand", productName, dateFrom, dateTo);

        // 2) Розничная продажа (retaildemand)
        total += calcSalesFromEntity("retaildemand", productName, dateFrom, dateTo);

        return total;
    }

    // Сканируем одну сущность (demand или retaildemand), без ограничения "розничный покупатель"
    private static double calcSalesFromEntity(String entityName,
                                              String productName,
                                              String dateFrom,
                                              String dateTo)
            throws IOException, InterruptedException, URISyntaxException {

        double totalQty = 0.0;

        System.out.println("=== Загрузка продаж из entity/" + entityName + " ===");

        int limit = 100;
        int offset = 0;
        int page = 0;
        int totalDocs = 0;

        String rawFilter = "moment>=" + dateFrom +
                ";moment<=" + dateTo +
                ";applicable=true";

        String filter = URLEncoder.encode(rawFilter, StandardCharsets.UTF_8);

        while (true) {
            page++;

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

            System.out.println("  " + entityName + " – страница " + page + ": " + rows.size() + " документов");
            totalDocs += rows.size();

            for (JsonNode doc : rows) {
                String docHref = doc.path("meta").path("href").asText("");
                if (docHref.isEmpty()) continue;

                // просто считаем позиции, без проверки тегов контрагента
                double qtyForDoc = loadPositionsForDoc(docHref, productName);
                if (qtyForDoc > 0) {
                    totalQty += qtyForDoc;
                    // сохраняем для вывода ссылок
                    RETAIL_ORDERS.put(docHref, qtyForDoc);
                }
            }

            if (rows.size() < limit) {
                break;
            }

            offset += limit;
        }

        System.out.println("Всего обработано документов " + entityName + ": " + totalDocs);
        System.out.printf(Locale.US,
                "Товар \"%s\" продан по entity/%s суммарно: %.2f шт%n",
                productName, entityName, totalQty);

        return totalQty;
    }


    /**
     * Загружает позиции retaildemand-документа и возвращает количество
     * по строкам, где ассортимент = нужный товар (по имени).
     */
    private static double loadPositionsForDoc(String docHref,
                                              String productName)
            throws IOException, InterruptedException, URISyntaxException {

        double qtyForDoc = 0.0;

        int limit = 100;
        int offset = 0;

        while (true) {
            String url = docHref + "/positions?limit=" + limit + "&offset=" + offset;
            String json = get(url);

            JsonNode root;
            try {
                root = MAPPER.readTree(json);
            } catch (JsonProcessingException e) {
                System.out.println("  retaildemand positions: неожиданный ответ, не JSON. Текст:\n" + json);
                break;
            }

            JsonNode rows = root.path("rows");
            if (!rows.isArray() || rows.size() == 0) {
                break;
            }

            for (JsonNode position : rows) {
                JsonNode assortmentMeta = position.path("assortment").path("meta");
                String productHref = assortmentMeta.path("href").asText();
                if (productHref == null || productHref.isBlank()) {
                    continue;
                }

                double quantity = position.path("quantity").asDouble(0.0);

                // Получаем имя товара по href (с кешем)
                String name = resolveProductName(productHref);

                if (name != null && name.equalsIgnoreCase(productName)) {
                    qtyForDoc += quantity;
                }
            }

            if (rows.size() < limit) {
                break;
            }

            offset += limit;
        }

        return qtyForDoc;
    }

    /**
     * Определяем, является ли контрагент "розничным покупателем" по тегам.
     */
    private static boolean isRetailBuyerByTags(String counterpartyHref)
            throws IOException, InterruptedException, URISyntaxException {

        if (counterpartyHref == null || counterpartyHref.isBlank()) {
            return false;
        }

        if (RETAIL_BUYER_CACHE.containsKey(counterpartyHref)) {
            return RETAIL_BUYER_CACHE.get(counterpartyHref);
        }

        String json = get(counterpartyHref);
        JsonNode root = MAPPER.readTree(json);

        boolean isRetail = false;

        JsonNode tagsNode = root.path("tags");
        if (tagsNode.isArray()) {
            for (JsonNode t : tagsNode) {
                String tag = t.asText("").toLowerCase(Locale.ROOT);
                if (tag.contains("розничный покупатель")) {
                    isRetail = true;
                    break;
                }
            }
        }

        RETAIL_BUYER_CACHE.put(counterpartyHref, isRetail);
        return isRetail;
    }

    /**
     * Возвращает name товара по его href, используя кеш.
     */

    private static String resolveProductName(String productHref)
            throws IOException, InterruptedException, URISyntaxException {

        if (productHref == null || productHref.isBlank()) return null;

        if (PRODUCT_NAME_CACHE.containsKey(productHref)) {
            return PRODUCT_NAME_CACHE.get(productHref);
        }

        // БЕЗ ?fields=name — как в твоём рабочем коде
        String json = get(productHref);
        JsonNode node = MAPPER.readTree(json);
        String name = node.path("name").asText("");

        PRODUCT_NAME_CACHE.put(productHref, name);
        return name;
    }

    // =====================================================================
    // HTTP GET + разжатие gzip
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
                if (attempt == maxAttempts) {
                    throw e;
                }
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
                while ((line = br.readLine()) != null) {
                    sb.append(line);
                }
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

    /**
     * Превращаем API-href документа в веб-ссылку МойСклад.
     * Пример:
     *  https://api.moysklad.ru/api/remap/1.2/entity/retaildemand/UUID
     *  → https://online.moysklad.ru/app/#retaildemand/edit?id=UUID
     */
    private static String buildWebLinkFromApiHref(String docHref) {
        if (docHref == null || docHref.isBlank()) return "";
        int idx = docHref.lastIndexOf('/');
        if (idx == -1 || idx == docHref.length() - 1) return "";
        String id = docHref.substring(idx + 1);
        return "https://online.moysklad.ru/app/#retaildemand/edit?id=" + id;
    }
}
