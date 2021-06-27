import model.Book;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

public class Main {

    public static void main(String[] args) throws IOException {
        Scanner scanner = new Scanner(System.in);

        List<Book> bookList = new ArrayList<>();
        int bookCount;
        System.out.println("How many books want you add");
        bookCount = Integer.parseInt(scanner.nextLine());
        for (int i = 0; i < bookCount; i++) {
            Book book = new Book();
            System.out.println("please input book title");
            book.setTitle(scanner.nextLine());
            System.out.println("please input book description");
            book.setDescription(scanner.nextLine());
            System.out.println("please input book price");
            book.setPrice(Double.parseDouble(scanner.nextLine()));
            System.out.println("please input book count");
            book.setCount(Integer.parseInt(scanner.nextLine()));
            bookList.add(i, book);

        }

        bookToExcel(bookList);


    }

    private static void bookToExcel(List<Book> bookList) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheetBooks = workbook.createSheet("Books");
        Row row = sheetBooks.createRow(0);

        Cell cellTitle = row.createCell(0);
        cellTitle.setCellValue("title");

        Cell cellDescription = row.createCell(1);
        cellDescription.setCellValue("description");

        Cell cellPrice = row.createCell(2);
        cellPrice.setCellValue("price");

        Cell cellCount = row.createCell(3);
        cellCount.setCellValue("Count");
        int rowIndex = 1;
        for (Book book : bookList) {
            Row bookRow = sheetBooks.createRow(rowIndex++);

            Cell cellBookTitle = bookRow.createCell(0);
            cellBookTitle.setCellValue(book.getTitle());

            Cell cellBookDescription = bookRow.createCell(1);
            cellBookDescription.setCellValue(book.getDescription());

            Cell cellBookPrice = bookRow.createCell(2);
            cellBookPrice.setCellValue(book.getPrice());

            Cell cellBookCount = bookRow.createCell(3);
            cellBookCount.setCellValue(book.getCount());

        }
        FileOutputStream fileOutputStream = new FileOutputStream("C:\\Users\\User\\IdeaProjects\\bookProject\\src\\main\\resources\\Book.xlsx");
        workbook.write(fileOutputStream);
        workbook.close();


    }
}
