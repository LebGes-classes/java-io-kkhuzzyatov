import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.IOException;
import java.util.ArrayList;

public class Main {
    public static void main(String[] args) throws IOException, InvalidFormatException {
        ArrayList<Student> students = Student.readStudents();
        Student.encodeStudents(students);

        ArrayList<Educator> educators = Educator.readEducators();
        Educator.encodeEducators(educators);
    }
}
