package model;


import lombok.*;

@Data
@AllArgsConstructor
@NoArgsConstructor

public class Author {

    private String name;
    private String surname;
    private String email;
    private Gender gender;

}