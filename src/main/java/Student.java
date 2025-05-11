import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.json.JSONArray;
import org.json.JSONObject;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.file.Paths;
import java.util.ArrayList;

public class Student {
    private String id;
    private String fullName;
    private String group;
    private ArrayList<Mark> marks = new ArrayList<>();

    private static final String STUDENT_XLSX_FILE_PATH = "./files/excel_files/student.xlsx";

    public Student(String id, String fullName, String group, ArrayList<Mark> marks) {
        this.id = id;
        this.fullName = fullName;
        this.group = group;
        this.marks = marks;
    }

    public static ArrayList<Student> readStudents() throws IOException, InvalidFormatException {
        ArrayList<Student> students = new ArrayList<>();

        Workbook workbook = WorkbookFactory.create(new File(STUDENT_XLSX_FILE_PATH));

        // Получение листа по индексу (первый лист)
        Sheet sheet = workbook.getSheetAt(0);

        // Создание DataFormatter для форматирования и получения значения ячейки как String
        DataFormatter dataFormatter = new DataFormatter();

        // цикл for-each для итерации по строкам и столбцам
        for (Row row: sheet) {
            if (row.getRowNum() != 0) {
                String id = dataFormatter.formatCellValue(row.getCell(0));
                String fullName = dataFormatter.formatCellValue(row.getCell(1));
                String group = dataFormatter.formatCellValue(row.getCell(2));

                ArrayList<Mark> marks = Mark.readMarksOfStudent(id);

                Student student = new Student(id, fullName, group, marks);
                students.add(student);
            }
        }

        // Закрытие workbook
        workbook.close();

        return students;
    }

    public static void encodeStudents(ArrayList<Student> students) throws IOException, InvalidFormatException {
        JSONArray studentsJSON = new JSONArray();
        for (Student student : students) {
            JSONObject studentJSON = new JSONObject();
            studentJSON.put("id", student.getId());
            studentJSON.put("full_name", student.getFullName());
            studentJSON.put("group", student.getGroup());

            JSONArray marksJSON = new JSONArray();
            for (Mark mark : student.getMarks()) {
                JSONObject markJSON = new JSONObject();
                markJSON.put("id", mark.getId());
                markJSON.put("subject", Subject.getNameById(mark.getSubjectId()));
                markJSON.put("value", mark.getValue());

                marksJSON.put(markJSON);
            }

            studentJSON.put("marks", marksJSON);

            studentsJSON.put(studentJSON);
        }

        String currentDir = System.getProperty("user.dir") + "/files/json_files";
        String filePath = Paths.get(currentDir, "students.json").toString();

        try (FileWriter file = new FileWriter(filePath)) {
            file.write(studentsJSON.toString(4)); // 4 - отступ для красивого форматирования
            System.out.println("JSON успешно сохранен в: " + filePath);
        } catch (IOException e) {
            System.out.println("Ошибка при сохранении файла: " + e.getMessage());
        }
    }

    public String getId() {
        return id;
    }

    public String getFullName() {
        return fullName;
    }

    public String getGroup() {
        return group;
    }

    public ArrayList<Mark> getMarks() {
        return marks;
    }
}
