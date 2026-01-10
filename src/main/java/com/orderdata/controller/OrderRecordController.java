package com.orderdata.controller;

import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.validation.annotation.Validated;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.orderdata.dto.OrderRecordRequest;
import com.orderdata.dto.OrderRecordResponse;
import com.orderdata.service.OrderRecordService;

import jakarta.validation.Valid;

@RestController
@RequestMapping("/api/order-records")
@Validated
public class OrderRecordController {

    private final OrderRecordService orderRecordService;

    public OrderRecordController(OrderRecordService orderRecordService) {
        this.orderRecordService = orderRecordService;
    }

    @PostMapping
    public ResponseEntity<OrderRecordResponse> storeOrderRecord(@Valid @RequestBody OrderRecordRequest request) {
        OrderRecordResponse response = orderRecordService.storeOrderRecord(request);
        return ResponseEntity.status(HttpStatus.CREATED).body(response);
    }
}
