package storage;


import model.Book;

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


}








