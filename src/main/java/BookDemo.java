import command.Commands;
import exception.AuthorNotFoundException;
import model.Author;
import model.Book;
import model.User;
import model.UserType;
import storage.AuthorStorage;
import storage.BookStorage;
import storage.UserStorage;
import model.Gender;

import java.io.IOException;
import java.util.Scanner;

class BookDemo implements Commands {
    private static Scanner scanner = new Scanner(System.in);
    private static BookStorage bookStorage = new BookStorage();
    private static AuthorStorage authorStorage = new AuthorStorage();
    private static UserStorage userStorage = new UserStorage();


    private static User currentUser = null;


    public static void main(String[] args) {
        initData();
        boolean run = true;


        while (run) {
            Commands.printLoginCommands();
            int command;
            try {
                command = Integer.parseInt(scanner.nextLine());
            } catch (NumberFormatException e) {
                command = -1;

            }
            switch (command) {
                case EXIT:
                    run = false;
                    break;
                case LOGIN:
                    login();
                    break;
                case REGISTER:
                    register();
                    break;
                default:
                    System.out.println("Invalid command");
            }
        }
    }

    private static void login() {
        System.out.println("Please input email,password");
        String emailPasswordStr = scanner.nextLine();
        String[] emailPassword = emailPasswordStr.split(",");
        User user = userStorage.getUserByEmail(emailPassword[0]);
        if (user == null) {
            System.out.println("User with " + emailPassword[0] + " does not exist");
        } else {
            if (user.getPassword().equals(emailPassword[1])) {
                currentUser = user;
                if (user.getUserType() == UserType.ADMIN) {
                    loginAdmin();
                } else if (user.getUserType() == UserType.USER) {
                    loginUser();
                }

            } else {
                System.out.println("Password is wrong!");
            }
        }
    }

    private static void loginUser() {
        System.out.println("Welcome " + currentUser.getName());

        boolean run = true;
        while (run) {
            Commands.printUserCommands();
            int command;
            try {
                command = Integer.parseInt(scanner.nextLine());

            } catch (NumberFormatException e) {
                command = -1;
            }

            switch (command) {
                case LOGOUT:
                    run = false;
                    break;
                case ADD_BOOK:
                    addBook();
                    break;
                case PRINT_ALL_BOOKS:
                    bookStorage.print();
                    break;
                case PRINT_BOOKS_BY_AUTHOR_NAME:
                    printBooksByAuthorName();
                    break;
                case PRINT_BOOKS_BY_GENRE:
                    printBooksByGenre();
                    break;
                case PRINT_BOOKS_BY_PRICE_RANGE:
                    printBooksByPriceRange();
                    break;
                case DOWNLOAD_BOOK_EXCEL:
                    downloadBooksExcel();
                    break;
                default:
                    System.out.println("Invalid command!");


            }
        }
    }

    private static void downloadBooksExcel() {
        System.out.println("Please input file location");
        String fileDir = scanner.nextLine();
        try {
            bookStorage.writeStudentToExcel(fileDir);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void register() {
        System.out.println("Please input name,surname,email,password");
        String userDataStr = scanner.nextLine();
        String[] userData = userDataStr.split(",");
        if (userData.length < 4) {
            System.out.println("Please input correct data!");
        } else {
            if (userStorage.getUserByEmail(userData[0]) == null) {
                User user = new User();
                user.setName(userData[0]);
                user.setSurname(userData[1]);
                user.setEmail(userData[2]);
                user.setPassword(userData[3]);
                user.setUserType(UserType.USER);
                userStorage.add(user);
                System.out.println("User created");

            } else {
                System.out.println("User with " + userData[0] + " already exist!");
            }
        }


    }

    private static void loginAdmin() {
        System.out.println("Welcome " + currentUser.getName());


        boolean run = true;

        while (run) {
            Commands.printAdminCommands();
            int command;
            try {
                command = Integer.parseInt(scanner.nextLine());
            } catch (NumberFormatException e) {
                command = -1;

            }
            switch (command) {
                case LOGOUT:
                    run = false;
                    break;
                case ADD_BOOK:
                    addBook();
                    break;
                case PRINT_ALL_BOOKS:
                    bookStorage.print();
                    break;
                case PRINT_BOOKS_BY_AUTHOR_NAME:
                    printBooksByAuthorName();
                    break;
                case PRINT_BOOKS_BY_GENRE:
                    printBooksByGenre();
                    break;
                case PRINT_BOOKS_BY_PRICE_RANGE:
                    printBooksByPriceRange();
                    break;
                case ADD_AUTHOR:
                    addAuthor();
                    break;
                case PRINT_ALL_AUTHOR:
                    authorStorage.print();
                    break;
                case DOWNLOAD_BOOK_EXCEL:
                    downloadBooksExcel();
                    break;
                default:
                    System.out.println("Invalid command");
            }
        }


    }


    private static void initData() {
        User admin = new User("admin", "admin", "admin@mail.ru", "12345", UserType.ADMIN);
        userStorage.add(admin);
        Author herbertSchildt = new Author("Herbert", "Schildt", "Shcildt@mail.com", Gender.MALE);
        Author georgeMartin = new Author("George", "Martin", "Martin@mail.com", Gender.MALE);
        Author sevak = new Author("Պարույր", " Սևակ", "Sevak@mail.com", Gender.MALE);
        authorStorage.add(herbertSchildt);
        authorStorage.add(georgeMartin);
        authorStorage.add(sevak);
        bookStorage.add(new Book("java", herbertSchildt, 1, 5000, "educational", admin));
        bookStorage.add(new Book("Dance with Dragons", georgeMartin, 5, 3500, "historical", admin));
        bookStorage.add(new Book("Մարդ էլ կա, մարդ էլ", sevak, 10, 8000, "բանաստեղծություններ", admin));
    }


    private static void printBooksByPriceRange() {
        System.out.println("please input min and max prices books");
        try {
            double priceMin = Double.parseDouble(scanner.nextLine());
            double priceMax = Double.parseDouble(scanner.nextLine());
            bookStorage.printBooksByPriceRange(priceMin, priceMax);
        } catch (NumberFormatException e) {

            System.out.println("No books found in that price range please choose other prices");
            printBooksByPriceRange();
        }


    }

    private static void addBook() {
        if (authorStorage.getSize() != 0) {
            authorStorage.print();
            System.out.println("Please choose lesson index");
            try {

                int authorIndex = Integer.parseInt(scanner.nextLine());
                Author author = authorStorage.getAuthorByIndex(authorIndex);


                System.out.println("please choose author name");
                String authorName = scanner.nextLine();
                bookStorage.printBooksByAuthorName(authorName);

                System.out.println("Please input books name");
                String title = scanner.nextLine();
                System.out.println("Please input books count");
                int count = Integer.parseInt(scanner.nextLine());
                System.out.println("Please input books price");
                double price = Double.parseDouble(scanner.nextLine());
                System.out.println("please input the book genre");
                String genre = scanner.nextLine();


                Book book = new Book(title, author, count, price, genre, currentUser);
                bookStorage.add(book);
                System.out.println("Book created");


            } catch (AuthorNotFoundException | NumberFormatException e) {
                System.out.println("Please choose correct number or index");
                addBook();

            }
        }
    }


    private static void printBooksByAuthorName() {
        System.out.println("Please input author name");
        String authorName = scanner.nextLine();
        bookStorage.printBooksByAuthorName(authorName);
    }

    private static void printBooksByGenre() {
        System.out.println("Please input books genre");
        String bookName = scanner.nextLine();
        bookStorage.printBooksByGenre(bookName);
    }


    private static void addAuthor() {

        if (authorStorage.getSize() != 0) {
            authorStorage.print();
            System.out.println("Please input author name");
            String name = scanner.nextLine();
            System.out.println("Please input authors surname");
            String surname = scanner.nextLine();
            System.out.println("Please input authors email");
            String email = scanner.nextLine();
            System.out.println("Please input authors gender");
            Gender gender = null;
            Author author;


            try {
                gender = Gender.valueOf(scanner.nextLine().toUpperCase());
            } catch (IllegalArgumentException e) {
                System.out.println("Please choose correct gender!!!");
                addAuthor();
            }
            if (!email.contains("@")) {
                System.out.println("Please choose correct email");
                addAuthor();

            } else {


                author = new Author(name, surname, email, gender);
                authorStorage.add(author);
                System.out.println("Author created");
            }
        }
    }


}