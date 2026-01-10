package com.orderdata.exception;

public class OrderStorageException extends RuntimeException {

    public OrderStorageException(String message) {
        super(message);
    }

    public OrderStorageException(String message, Throwable cause) {
        super(message, cause);
    }
}
