package com.ilivoo;

import java.io.IOException;
import java.nio.file.FileVisitResult;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.SimpleFileVisitor;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.ArrayList;
import java.util.List;

public class FileUtil {

    public static List<String> listFiles(String... filePaths) throws IOException {
        List<String> result = new ArrayList<>();
        for (String filePath : filePaths) {
            Path path = Paths.get(filePath);
            if (Files.isRegularFile(path)) {
                result.add(filePath);
            } else {
                Files.walkFileTree(path, new SimpleFileVisitor<Path>() {
                    @Override
                    public FileVisitResult visitFile(Path file, BasicFileAttributes attrs) throws IOException {
                        result.add(file.toString());
                        return super.visitFile(file, attrs);
                    }
                });
            }
        }
        return result;
    }

    public static boolean ensureNoFile(String filePath) throws IOException {
        Path path = Paths.get(filePath);
        if (Files.exists(path) && Files.isRegularFile(path)) {
            Files.delete(path);
        } else {
            throw new RuntimeException(filePath + ": is not regular file");
        }
        return true;
    }
}
