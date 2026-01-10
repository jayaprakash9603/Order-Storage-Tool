package com.orderdata.dto;

public class OrderRecordResponse {

    private final String message;
    private final String filePath;
    private final String orderId;
    private final String storedAt;

    public OrderRecordResponse(String message, String filePath, String orderId, String storedAt) {
        this.message = message;
        this.filePath = filePath;
        this.orderId = orderId;
        this.storedAt = storedAt;
    }

    public String getMessage() {
        return message;
    }

    public String getFilePath() {
        return filePath;
    }

    public String getOrderId() {
        return orderId;
    }

    public String getStoredAt() {
        return storedAt;
    }
}
