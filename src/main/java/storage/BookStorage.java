package storage;


import model.Book;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class BookStorage {

    private Book[] array = new Book[10];
    private int size;

    public Book add(Book book) {
        if (size == array.length) {
            extendArray();

        }
        array[size++] = book;
        return book;
    }

    private void extendArray() {
        Book[] temp = new Book[array.length + 10];
        System.arraycopy(array, 0, temp, 0, size);
        array = temp;

    }

    public void print() {
        for (int i = 0; i < size; i++) {
            System.out.println(i + "." + array[i]);


        }

    }


    public void printBooksByAuthorName(String authorName) {
        for (int i = 0; i < size; i++) {
            if (array[i].getAuthor().getName().equals(authorName)) {
                System.out.println(array[i]);
            }
        }


    }

    public void printBooksByGenre(String bookGenre) {
        for (int i = 0; i < size; i++) {
            if (array[i].getGenre().equals(bookGenre)) {
                System.out.println(array[i]);
            }


        }
    }


    public void printBooksByPriceRange(double priceMin, double priceMax) {
        for (int i = 0; i < size; i++) {
            if (array[i].getPrice() >= priceMin && array[i].getPrice() <= priceMax) {


                System.out.println(array[i]);


            }


        }

    }

    public int getSize() {
        return size;
    }


    public void writeStudentToExcel(String fileDir) throws IOException {
        File directory = new File(fileDir);
        if (directory.isFile()) {
            throw new RuntimeException("fileDir must be a directory");
        }
        File excelFile = new File(directory, "books_" + System.currentTimeMillis() + ".xlsx");
        excelFile.createNewFile();
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("books");
        Row headerRow = sheet.createRow(0);

        Cell nameCell = headerRow.createCell(0);
        nameCell.setCellValue("Title");

        Cell surname = headerRow.createCell(1);
        surname.setCellValue("Count");

        Cell age = headerRow.createCell(2);
        age.setCellValue("Genre");

        Cell phoneCall = headerRow.createCell(3);
        phoneCall.setCellValue("Price");

        for (int i = 0; i < size; i++) {
            Book book = array[i];
            Row row = sheet.createRow(i + 1);
            row.createCell(0).setCellValue(book.getTitle());
            row.createCell(1).setCellValue(book.getCount());
            row.createCell(2).setCellValue(book.getGenre());
            row.createCell(3).setCellValue(book.getPrice());

        }

        workbook.write(new FileOutputStream(excelFile));
        System.out.println("Excel was created successfully");

    }
}








