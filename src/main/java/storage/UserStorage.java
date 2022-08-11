package storage;


import model.User;

public class UserStorage {

    private User[] array = new User[10];
    private int size;

    public User add(User book) {
        if (size == array.length) {
            extendArray();

        }
        array[size++] = book;
        return book;
    }

    private void extendArray() {
        User[] temp = new User[array.length + 10];
        System.arraycopy(array, 0, temp, 0, size);
        array = temp;

    }

    public void print() {
        for (int i = 0; i < size; i++) {
            System.out.println(i + "." + array[i]);


        }

    }
    public User getUserByEmail(String email) {
        for (int i = 0; i < size; i++) {
            if (array[i].getEmail().equals(email)) {
                return array[i];

            }
        }
        return null;
    }

}







