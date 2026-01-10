package com.orderdata.dto;

import jakarta.validation.constraints.NotBlank;
import jakarta.validation.constraints.Pattern;

public class OrderRecordRequest {

    @NotBlank(message = "orderId is required")
    private String orderId;

    @NotBlank(message = "directoryPath is required")
    private String directoryPath;

    @NotBlank(message = "env is required")
    private String env;

    @Pattern(regexp = "^$|^\\d{4}-\\d{2}-\\d{2}$", message = "date must be in yyyy-MM-dd format")
    private String date;

    public String getOrderId() {
        return orderId;
    }

    public void setOrderId(String orderId) {
        this.orderId = orderId;
    }

    public String getDirectoryPath() {
        return directoryPath;
    }

    public void setDirectoryPath(String directoryPath) {
        this.directoryPath = directoryPath;
    }

    public String getEnv() {
        return env;
    }

    public void setEnv(String env) {
        this.env = env;
    }

    public String getDate() {
        return date;
    }

    public void setDate(String date) {
        this.date = date;
    }
}
