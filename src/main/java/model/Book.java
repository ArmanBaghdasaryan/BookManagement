package model;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@AllArgsConstructor
@NoArgsConstructor


public class Book {

    private String title;
    private Author author;
    private int count;
    private double price;
    private String genre;
    private User registeredUser;

}