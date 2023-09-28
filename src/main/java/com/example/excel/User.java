package com.example.excel;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
@AllArgsConstructor
public class User {
    private String firstName, lastName;
    private int order_quantity;
    private double total_spent;
}
