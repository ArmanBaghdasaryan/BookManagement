package storage;


import exception.AuthorNotFoundException;
import model.Author;

public class AuthorStorage {

    private Author[] array = new Author[10];
    private int size;

    public Author add(Author author) {
        if (size == array.length) {
            extendArray();

        }
        array[size++] = author;
        return author;
    }

    private void extendArray() {
        Author[] temp = new Author[array.length + 10];
        for (int i = 0; i < size; i++) {
            temp[i] = array[i];

        }
        array = temp;

    }

    public void print() {
        for (int i = 0; i < size; i++) {
            System.out.println(i + "." + array[i]);


        }
    }

    public int getSize() {
        return size;
    }


    public Author getAuthorByIndex(int index) throws AuthorNotFoundException {


        if (index < 0 || index >= size) {
            throw new AuthorNotFoundException("Author with " + index + " does not exist");

        }
        return array[index];
    }


}















