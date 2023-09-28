package com.example.excel;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;


@Getter
@Setter
@AllArgsConstructor
public class Producto {
    
    private String name, category;
    private double price;
    private Boolean active;
    private int quantity_sold;
}
