package com.orderdata.service;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.time.temporal.ChronoUnit;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.NavigableMap;
import java.util.TreeMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import com.orderdata.dto.OrderRecordRequest;
import com.orderdata.dto.OrderRecordResponse;
import com.orderdata.exception.OrderStorageException;

@Service
public class OrderRecordService {

    private static final String FILE_NAME = "order-records.xlsx";
    private static final String ORDERS_SHEET_NAME = "Orders";
    private static final String ALL_ORDERS_SHEET_NAME = "AllOrders";
    private static final String VALUE_CONDITIONS_SHEET_NAME = "ValueConditions";
    private static final DateTimeFormatter DATE_ONLY_FORMATTER = DateTimeFormatter.ofPattern("yyyy-MM-dd");
    private static final String ORDER_ID_HEADER_LABEL = "Order ID";
    private static final String PLAN_HEADER_LABEL = "Plan Name";
    private static final String ENVIRONMENT_HEADER_LABEL = "Environment";
    private static final String SPRINT_HEADER_LABEL = "Sprint";
    private static final String DATE_HEADER_LABEL = "Date";
    private static final int COLUMN_OFFSET = 3;
    private static final int TOP_PADDING_ROWS = 3;
    private static final int TABLE_GAP_ROWS = 2;
    private static final int FIRST_SPRINT_NUMBER = 340;
    private static final int DEFAULT_LAST_SPRINT_NUMBER = 398;
    private static final int ANCHOR_SPRINT_NUMBER = 385;
    private static final LocalDate ANCHOR_SPRINT_START = LocalDate.of(2026, 1, 7);
    private static final int SPRINT_INTERVAL_DAYS = 14;
    private static final int SPRINT_WINDOW_DAYS = 14;
    private static final int ADDITIONAL_SPRINTS = 15;
    private static final int VALUE_CONDITIONS_SHEET_INDEX = 2;
    private static final int VALUE_CONDITIONS_TOP_PADDING = 2;
    private static final int VALUE_CONDITIONS_COLUMN_OFFSET = 2;
    private static final Pattern PLAN_PATTERN = Pattern.compile("^(NEW|MOD|CAN)-(\\d+)", Pattern.CASE_INSENSITIVE);
    private static final Pattern UUID_TIMESTAMP_PATTERN = Pattern.compile(
            "^[0-9a-fA-F-]{8,}/\\d{4}-\\d{2}-\\d{2}T\\d{2}:\\d{2}:\\d{2}(?:\\.\\d+)?Z$",
            Pattern.CASE_INSENSITIVE);
    private static final Pattern SPRINT_LABEL_PATTERN = Pattern.compile("^Sprint-\\d+$", Pattern.CASE_INSENSITIVE);

    public OrderRecordResponse storeOrderRecord(OrderRecordRequest request) {
        Path directoryPath = resolveDirectory(request.getDirectoryPath());
        Path filePath = directoryPath.resolve(FILE_NAME);
        boolean fileExists = Files.exists(filePath);

        Workbook workbook = fileExists ? openExistingWorkbook(filePath) : new XSSFWorkbook();
        try (workbook) {
            NavigableMap<LocalDate, List<OrderEntry>> existingEntries = readExistingEntries(workbook);

            LocalDate orderDate = resolveOrderDate(request.getDate());
            String environment = normalizeEnvironment(request.getEnv());
            existingEntries
                    .computeIfAbsent(orderDate, date -> new ArrayList<>())
                    .add(new OrderEntry(request.getOrderId(), environment));

            CellStyle headerStyle = createHeaderCellStyle(workbook);
            LocalDate latestRelevantDate = determineLatestRelevantDate(existingEntries, orderDate);
            List<SprintWindow> sprintWindows = ensureSprintWindows(workbook, latestRelevantDate, headerStyle);
            rebuildOrdersSheet(workbook, existingEntries, headerStyle, sprintWindows);
            rebuildAllOrdersSheet(workbook, existingEntries, headerStyle, sprintWindows);
            positionValueConditionsSheet(workbook);

            writeWorkbookToFile(workbook, filePath);

            return new OrderRecordResponse(
                    "Order ID saved successfully",
                    filePath.toString(),
                    request.getOrderId(),
                    orderDate.format(DATE_ONLY_FORMATTER));
        } catch (IOException ioException) {
            throw new OrderStorageException("Failed to store order details", ioException);
        }
    }

    private Path resolveDirectory(String directory) {
        try {
            Path path = Paths.get(directory).toAbsolutePath().normalize();
            Files.createDirectories(path);
            return path;
        } catch (IOException ioException) {
            throw new OrderStorageException("Unable to create or access directory: " + directory, ioException);
        }
    }

    private Workbook openExistingWorkbook(Path filePath) {
        try (InputStream inputStream = Files.newInputStream(filePath)) {
            return new XSSFWorkbook(inputStream);
        } catch (IOException ioException) {
            throw new OrderStorageException("Unable to read existing Excel file: " + filePath, ioException);
        }
    }

    private void autoSizeColumns(Sheet sheet, int startColumnIndex) {
        for (int columnIndex = startColumnIndex; columnIndex < startColumnIndex + 5; columnIndex++) {
            sheet.autoSizeColumn(columnIndex);
        }
    }

    private void writeWorkbookToFile(Workbook workbook, Path filePath) throws IOException {
        try (OutputStream outputStream = Files.newOutputStream(filePath)) {
            workbook.write(outputStream);
        }
    }

    private CellStyle createHeaderCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        return style;
    }

    private String getStringValue(Cell cell) {
        if (cell == null) {
            return null;
        }

        if (cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue();
        }

        return cell.toString();
    }

    private String firstNonBlank(String... candidates) {
        if (candidates == null) {
            return null;
        }
        for (String candidate : candidates) {
            if (candidate != null && !candidate.trim().isEmpty()) {
                return candidate;
            }
        }
        return null;
    }

    private boolean equalsIgnoreCase(String expected, String actual) {
        return expected != null && actual != null && expected.equalsIgnoreCase(actual.trim());
    }

    private boolean matchesNewOrderHeader(String orderCell, String planCell, String environmentCell, String sprintCell,
            String dateCell) {
        return equalsIgnoreCase(ORDER_ID_HEADER_LABEL, orderCell)
                && equalsIgnoreCase(PLAN_HEADER_LABEL, planCell)
                && equalsIgnoreCase(ENVIRONMENT_HEADER_LABEL, environmentCell)
                && equalsIgnoreCase(SPRINT_HEADER_LABEL, sprintCell)
                && equalsIgnoreCase(DATE_HEADER_LABEL, dateCell);
    }

    private boolean matchesLegacyOrderHeader(String orderCell, String planCell, String sprintCell, String dateCell) {
        return equalsIgnoreCase(ORDER_ID_HEADER_LABEL, orderCell)
                && equalsIgnoreCase(PLAN_HEADER_LABEL, planCell)
                && equalsIgnoreCase(SPRINT_HEADER_LABEL, sprintCell)
                && equalsIgnoreCase(DATE_HEADER_LABEL, dateCell);
    }

    private LocalDate resolveOrderDate(String requestedDate) {
        if (requestedDate == null || requestedDate.isBlank()) {
            return LocalDate.now();
        }

        try {
            return LocalDate.parse(requestedDate.trim(), DATE_ONLY_FORMATTER);
        } catch (DateTimeParseException exception) {
            throw new OrderStorageException("Invalid date format. Use yyyy-MM-dd", exception);
        }
    }

    private LocalDate determineLatestRelevantDate(NavigableMap<LocalDate, List<OrderEntry>> entries,
            LocalDate fallback) {
        LocalDate latestEntryDate = entries.isEmpty() ? fallback : entries.lastKey();
        LocalDate today = LocalDate.now();
        return latestEntryDate.isAfter(today) ? latestEntryDate : today;
    }

    private NavigableMap<LocalDate, List<OrderEntry>> readExistingEntries(Workbook workbook) {
        NavigableMap<LocalDate, List<OrderEntry>> entries = new TreeMap<>();
        Sheet sourceSheet = workbook.getSheet(ORDERS_SHEET_NAME);
        if (sourceSheet == null) {
            sourceSheet = workbook.getSheet(ALL_ORDERS_SHEET_NAME);
        }
        if (sourceSheet == null) {
            return entries;
        }

        for (Row row : sourceSheet) {
            if (isOrderHeaderRow(row)) {
                continue;
            }

            String orderIdValue = firstNonBlank(
                    getStringValue(row.getCell(COLUMN_OFFSET)),
                    getStringValue(row.getCell(0)));

            String dateValue = resolveDateValue(row);
            String environmentValue = resolveEnvironmentValue(row);

            if (orderIdValue == null || dateValue == null) {
                continue;
            }

            try {
                LocalDate date = LocalDate.parse(dateValue.trim(), DATE_ONLY_FORMATTER);
                String trimmedOrderId = orderIdValue.trim();
                if (trimmedOrderId.isEmpty()) {
                    continue;
                }
                entries
                        .computeIfAbsent(date, key -> new ArrayList<>())
                        .add(new OrderEntry(trimmedOrderId, environmentValue));
            } catch (DateTimeParseException ignored) {
                // skip rows with non-date values
            }
        }
        return entries;
    }

    private void rebuildOrdersSheet(Workbook workbook, NavigableMap<LocalDate, List<OrderEntry>> entries,
            CellStyle headerStyle, List<SprintWindow> sprintWindows) {
        Sheet sheet = recreateSheet(workbook, ORDERS_SHEET_NAME);
        rebuildSheetWithEntries(sheet, entries, headerStyle, sprintWindows);
        autoSizeColumns(sheet, COLUMN_OFFSET);
    }

    private void rebuildAllOrdersSheet(Workbook workbook, NavigableMap<LocalDate, List<OrderEntry>> entries,
            CellStyle headerStyle, List<SprintWindow> sprintWindows) {
        Sheet sheet = recreateSheet(workbook, ALL_ORDERS_SHEET_NAME);
        int currentRowIndex = 0;
        Row headerRow = sheet.createRow(currentRowIndex++);
        populateHeaderRow(headerRow, headerStyle, 0);

        for (Map.Entry<LocalDate, List<OrderEntry>> entry : entries.entrySet()) {
            String sprintName = resolveSprintName(entry.getKey(), sprintWindows);
            for (OrderEntry orderEntry : entry.getValue()) {
                Row dataRow = sheet.createRow(currentRowIndex++);
                populateDataRow(dataRow, orderEntry, entry.getKey(), sprintName, 0);
            }
        }

        autoSizeColumns(sheet, 0);
    }

    private void positionValueConditionsSheet(Workbook workbook) {
        int sheetIndex = workbook.getSheetIndex(VALUE_CONDITIONS_SHEET_NAME);
        if (sheetIndex < 0) {
            return;
        }
        int targetIndex = Math.min(VALUE_CONDITIONS_SHEET_INDEX, workbook.getNumberOfSheets() - 1);
        workbook.setSheetOrder(VALUE_CONDITIONS_SHEET_NAME, targetIndex);
    }

    private Sheet recreateSheet(Workbook workbook, String sheetName) {
        int existingIndex = workbook.getSheetIndex(sheetName);
        if (existingIndex >= 0) {
            workbook.removeSheetAt(existingIndex);
        }
        return workbook.createSheet(sheetName);
    }

    private void autoSizeValueConditionsColumns(Sheet sheet, int columnOffset) {
        for (int i = 0; i < 5; i++) {
            sheet.autoSizeColumn(columnOffset + i);
        }
    }

    private void rebuildSheetWithEntries(Sheet sheet, NavigableMap<LocalDate, List<OrderEntry>> entries,
            CellStyle headerStyle, List<SprintWindow> sprintWindows) {
        ensureTopPadding(sheet);
        int currentRowIndex = TOP_PADDING_ROWS;

        for (Map.Entry<LocalDate, List<OrderEntry>> entry : entries.entrySet()) {
            Row headerRow = sheet.createRow(currentRowIndex++);
            populateHeaderRow(headerRow, headerStyle, COLUMN_OFFSET);

            String sprintName = resolveSprintName(entry.getKey(), sprintWindows);
            for (OrderEntry orderEntry : entry.getValue()) {
                Row dataRow = sheet.createRow(currentRowIndex++);
                populateDataRow(dataRow, orderEntry, entry.getKey(), sprintName, COLUMN_OFFSET);
            }

            currentRowIndex = addTableGap(sheet, currentRowIndex);
        }
    }

    private void ensureTopPadding(Sheet sheet) {
        for (int rowIndex = 0; rowIndex < TOP_PADDING_ROWS; rowIndex++) {
            sheet.createRow(rowIndex);
        }
    }

    private int addTableGap(Sheet sheet, int startRowIndex) {
        for (int i = 0; i < TABLE_GAP_ROWS; i++) {
            sheet.createRow(startRowIndex++);
        }
        return startRowIndex;
    }

    private void applyValueConditionHeaderCell(Row row, int columnIndex, String label, CellStyle headerStyle) {
        Cell cell = row.createCell(columnIndex);
        cell.setCellValue(label);
        cell.setCellStyle(headerStyle);
    }

    private void populateHeaderRow(Row headerRow, CellStyle headerStyle, int columnOffset) {
        Cell orderIdHeader = headerRow.createCell(columnOffset);
        orderIdHeader.setCellValue(ORDER_ID_HEADER_LABEL);
        orderIdHeader.setCellStyle(headerStyle);

        Cell planHeader = headerRow.createCell(columnOffset + 1);
        planHeader.setCellValue(PLAN_HEADER_LABEL);
        planHeader.setCellStyle(headerStyle);

        Cell envHeader = headerRow.createCell(columnOffset + 2);
        envHeader.setCellValue(ENVIRONMENT_HEADER_LABEL);
        envHeader.setCellStyle(headerStyle);

        Cell sprintHeader = headerRow.createCell(columnOffset + 3);
        sprintHeader.setCellValue(SPRINT_HEADER_LABEL);
        sprintHeader.setCellStyle(headerStyle);

        Cell dateHeader = headerRow.createCell(columnOffset + 4);
        dateHeader.setCellValue(DATE_HEADER_LABEL);
        dateHeader.setCellStyle(headerStyle);
    }

    private void populateDataRow(Row dataRow, OrderEntry entry, LocalDate date, String sprintName, int columnOffset) {
        Cell orderIdCell = dataRow.createCell(columnOffset);
        orderIdCell.setCellValue(entry.getOrderId());

        Cell planCell = dataRow.createCell(columnOffset + 1);
        planCell.setCellValue(determinePlan(entry.getOrderId()));

        Cell environmentCell = dataRow.createCell(columnOffset + 2);
        environmentCell.setCellValue(entry.getEnvironment());

        Cell sprintCell = dataRow.createCell(columnOffset + 3);
        sprintCell.setCellValue(sprintName);

        Cell dateCell = dataRow.createCell(columnOffset + 4);
        dateCell.setCellValue(date.format(DATE_ONLY_FORMATTER));
    }

    private String resolveDateValue(Row row) {
        int[] candidateColumns = new int[] {
                COLUMN_OFFSET + 4,
                COLUMN_OFFSET + 3,
                COLUMN_OFFSET + 2,
                4,
                3,
                2,
                1
        };

        for (int columnIndex : candidateColumns) {
            if (columnIndex < 0) {
                continue;
            }
            String rawValue = getStringValue(row.getCell(columnIndex));
            if (rawValue == null) {
                continue;
            }
            String trimmed = rawValue.trim();
            if (trimmed.isEmpty()) {
                continue;
            }
            if (isDateValue(trimmed)) {
                return trimmed;
            }
        }
        return null;
    }

    private String resolveEnvironmentValue(Row row) {
        String environmentValue = firstNonBlank(
                getStringValue(row.getCell(COLUMN_OFFSET + 2)),
                getStringValue(row.getCell(2)));
        String normalized = normalizeEnvironment(environmentValue);
        if (normalized.isEmpty()) {
            return "";
        }
        if (isSprintLabel(normalized) || isDateValue(normalized)) {
            return "";
        }
        return normalized;
    }

    private boolean isDateValue(String value) {
        try {
            LocalDate.parse(value, DATE_ONLY_FORMATTER);
            return true;
        } catch (DateTimeParseException ignored) {
            return false;
        }
    }

    private String normalizeEnvironment(String environment) {
        if (environment == null) {
            return "";
        }

        String trimmed = environment.trim();
        return trimmed;
    }

    private boolean isSprintLabel(String value) {
        return value != null && SPRINT_LABEL_PATTERN.matcher(value).matches();
    }

    private boolean isOrderHeaderRow(Row row) {
        if (row == null) {
            return false;
        }

        String primaryOrderHeader = getStringValue(row.getCell(COLUMN_OFFSET));
        String primaryPlanHeader = getStringValue(row.getCell(COLUMN_OFFSET + 1));
        String primaryEnvironmentHeader = getStringValue(row.getCell(COLUMN_OFFSET + 2));
        String primarySprintHeader = getStringValue(row.getCell(COLUMN_OFFSET + 3));
        String primaryDateHeader = getStringValue(row.getCell(COLUMN_OFFSET + 4));

        if (matchesNewOrderHeader(primaryOrderHeader, primaryPlanHeader, primaryEnvironmentHeader, primarySprintHeader,
                primaryDateHeader)) {
            return true;
        }

        if (matchesLegacyOrderHeader(primaryOrderHeader, primaryPlanHeader, primaryEnvironmentHeader,
                primarySprintHeader)) {
            return true;
        }

        String zeroOrderHeader = getStringValue(row.getCell(0));
        String zeroPlanHeader = getStringValue(row.getCell(1));
        String zeroEnvironmentHeader = getStringValue(row.getCell(2));
        String zeroSprintHeader = getStringValue(row.getCell(3));
        String zeroDateHeader = getStringValue(row.getCell(4));

        if (matchesNewOrderHeader(zeroOrderHeader, zeroPlanHeader, zeroEnvironmentHeader, zeroSprintHeader,
                zeroDateHeader)) {
            return true;
        }

        if (matchesLegacyOrderHeader(zeroOrderHeader, zeroPlanHeader, zeroEnvironmentHeader, zeroSprintHeader)) {
            return true;
        }

        String legacyDateHeader = getStringValue(row.getCell(1));
        return equalsIgnoreCase(ORDER_ID_HEADER_LABEL, zeroOrderHeader)
                && equalsIgnoreCase(DATE_HEADER_LABEL, legacyDateHeader);
    }

    private String determinePlan(String orderId) {
        if (orderId == null) {
            return "";
        }

        String trimmedId = orderId.trim();
        if (trimmedId.isEmpty()) {
            return "";
        }

        if (UUID_TIMESTAMP_PATTERN.matcher(trimmedId).matches()) {
            return "UCDM";
        }

        String upper = trimmedId.toUpperCase();

        if (upper.startsWith("TOM") || upper.startsWith("PRJ")) {
            return "PPM";
        }

        if (upper.startsWith("AMEX")) {
            return "NEW-14";
        }

        if (upper.startsWith("UCP")) {
            return "UCSS";
        }

        Matcher matcher = PLAN_PATTERN.matcher(trimmedId);
        if (matcher.find()) {
            return matcher.group(0).toUpperCase();
        }

        if (upper.startsWith("NEW")) {
            return "NEW";
        }
        if (upper.startsWith("MOD")) {
            return "MOD";
        }
        if (upper.startsWith("CAN")) {
            return "CAN";
        }
        return "";
    }

    private List<SprintWindow> ensureSprintWindows(Workbook workbook, LocalDate coverageDate, CellStyle headerStyle) {
        int requiredLastSprintNumber = determineRequiredLastSprintNumber(coverageDate);
        Sheet sheet = workbook.getSheet(VALUE_CONDITIONS_SHEET_NAME);

        List<SprintWindow> windows = sheet == null ? new ArrayList<>() : loadSprintWindows(sheet);
        int existingLastSprint = getLastSprintNumber(windows);

        if (sheet == null || windows.isEmpty() || existingLastSprint < requiredLastSprintNumber) {
            sheet = recreateValueConditionsSheet(workbook, requiredLastSprintNumber, headerStyle);
            windows = loadSprintWindows(sheet);
        }

        return windows;
    }

    private Sheet recreateValueConditionsSheet(Workbook workbook, int requiredLastSprintNumber, CellStyle headerStyle) {
        int existingIndex = workbook.getSheetIndex(VALUE_CONDITIONS_SHEET_NAME);
        if (existingIndex >= 0) {
            workbook.removeSheetAt(existingIndex);
        }
        Sheet sheet = workbook.createSheet(VALUE_CONDITIONS_SHEET_NAME);

        int currentRowIndex = VALUE_CONDITIONS_TOP_PADDING;
        int columnOffset = VALUE_CONDITIONS_COLUMN_OFFSET;

        Row infoRow = sheet.createRow(currentRowIndex++);
        infoRow.createCell(columnOffset).setCellValue("Current Sprint");
        infoRow.createCell(columnOffset + 1).setCellValue("Sprint-" + ANCHOR_SPRINT_NUMBER);
        infoRow.createCell(columnOffset + 2)
                .setCellValue("Start Date: " + ANCHOR_SPRINT_START.format(DATE_ONLY_FORMATTER));
        infoRow.createCell(columnOffset + 3)
                .setCellValue("End Date: " + calculateSprintEnd(ANCHOR_SPRINT_START).format(DATE_ONLY_FORMATTER));
        infoRow.createCell(columnOffset + 4).setCellValue("Window (days): " + SPRINT_WINDOW_DAYS);

        currentRowIndex++; // spacer row for visual separation

        Row headerRow = sheet.createRow(currentRowIndex++);
        applyValueConditionHeaderCell(headerRow, columnOffset, "Sprint Name", headerStyle);
        applyValueConditionHeaderCell(headerRow, columnOffset + 1, "Start Date", headerStyle);
        applyValueConditionHeaderCell(headerRow, columnOffset + 2, "End Date", headerStyle);

        int rowIndex = currentRowIndex;
        int lastSprintNumber = Math.max(requiredLastSprintNumber, DEFAULT_LAST_SPRINT_NUMBER);
        for (int sprintNumber = FIRST_SPRINT_NUMBER; sprintNumber <= lastSprintNumber; sprintNumber++) {
            LocalDate startDate = calculateSprintStart(sprintNumber);
            LocalDate endDate = calculateSprintEnd(startDate);
            Row dataRow = sheet.createRow(rowIndex++);
            dataRow.createCell(columnOffset).setCellValue("Sprint-" + sprintNumber);
            dataRow.createCell(columnOffset + 1).setCellValue(startDate.format(DATE_ONLY_FORMATTER));
            dataRow.createCell(columnOffset + 2).setCellValue(endDate.format(DATE_ONLY_FORMATTER));
        }

        autoSizeValueConditionsColumns(sheet, columnOffset);

        return sheet;
    }

    private int determineRequiredLastSprintNumber(LocalDate coverageDate) {
        int coverageSprintNumber = calculateSprintNumberForDate(coverageDate);
        int baselineSprint = Math.max(coverageSprintNumber, FIRST_SPRINT_NUMBER);
        return baselineSprint + ADDITIONAL_SPRINTS;
    }

    private int getLastSprintNumber(List<SprintWindow> windows) {
        if (windows == null || windows.isEmpty()) {
            return -1;
        }
        SprintWindow lastWindow = windows.get(windows.size() - 1);
        return parseSprintNumber(lastWindow.getName());
    }

    private int parseSprintNumber(String sprintName) {
        if (sprintName == null || !sprintName.startsWith("Sprint-")) {
            return -1;
        }
        try {
            return Integer.parseInt(sprintName.substring("Sprint-".length()));
        } catch (NumberFormatException exception) {
            return -1;
        }
    }

    private List<SprintWindow> loadSprintWindows(Sheet sheet) {
        List<SprintWindow> windows = new ArrayList<>();
        boolean headerReached = false;
        int columnOffset = VALUE_CONDITIONS_COLUMN_OFFSET;

        for (Row row : sheet) {
            String sprintCell = firstNonBlank(
                    getStringValue(row.getCell(columnOffset)),
                    getStringValue(row.getCell(0)));
            String startCell = firstNonBlank(
                    getStringValue(row.getCell(columnOffset + 1)),
                    getStringValue(row.getCell(1)));
            String endCell = firstNonBlank(
                    getStringValue(row.getCell(columnOffset + 2)),
                    getStringValue(row.getCell(2)));

            if (!headerReached) {
                if (equalsIgnoreCase("Sprint Name", sprintCell)
                        && equalsIgnoreCase("Start Date", startCell)
                        && equalsIgnoreCase("End Date", endCell)) {
                    headerReached = true;
                }
                continue;
            }

            if (sprintCell == null || sprintCell.trim().isEmpty() || startCell == null || endCell == null) {
                continue;
            }

            try {
                LocalDate startDate = LocalDate.parse(startCell.trim(), DATE_ONLY_FORMATTER);
                LocalDate endDate = LocalDate.parse(endCell.trim(), DATE_ONLY_FORMATTER);
                windows.add(new SprintWindow(sprintCell.trim(), startDate, endDate));
            } catch (Exception ignored) {
                // Skip invalid sprint definitions
            }
        }

        return windows;
    }

    private String resolveSprintName(LocalDate date, List<SprintWindow> sprintWindows) {
        if (sprintWindows == null || sprintWindows.isEmpty()) {
            return "";
        }

        for (SprintWindow window : sprintWindows) {
            if ((date.isEqual(window.getStart()) || date.isAfter(window.getStart()))
                    && (date.isEqual(window.getEnd()) || date.isBefore(window.getEnd()))) {
                return window.getName();
            }
        }
        return "";
    }

    private LocalDate calculateSprintStart(int sprintNumber) {
        long offset = (long) (sprintNumber - ANCHOR_SPRINT_NUMBER) * SPRINT_INTERVAL_DAYS;
        return ANCHOR_SPRINT_START.plusDays(offset);
    }

    private int calculateSprintNumberForDate(LocalDate date) {
        long daysBetween = ChronoUnit.DAYS.between(ANCHOR_SPRINT_START, date);
        long offset = Math.floorDiv(daysBetween, SPRINT_INTERVAL_DAYS);
        return (int) (ANCHOR_SPRINT_NUMBER + offset);
    }

    private LocalDate calculateSprintEnd(LocalDate startDate) {
        return startDate.plusDays(SPRINT_WINDOW_DAYS - 1);
    }

    private static final class OrderEntry {
        private final String orderId;
        private final String environment;

        private OrderEntry(String orderId, String environment) {
            this.orderId = orderId;
            this.environment = environment == null ? "" : environment;
        }

        private String getOrderId() {
            return orderId;
        }

        private String getEnvironment() {
            return environment;
        }
    }

    private static final class SprintWindow {
        private final String name;
        private final LocalDate start;
        private final LocalDate end;

        private SprintWindow(String name, LocalDate start, LocalDate end) {
            this.name = name;
            this.start = start;
            this.end = end;
        }

        private String getName() {
            return name;
        }

        private LocalDate getStart() {
            return start;
        }

        private LocalDate getEnd() {
            return end;
        }
    }
}
