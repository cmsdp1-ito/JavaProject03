package com.example;

import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.poifs.crypt.*;
import java.io.*;

public class App {
    public static void main(String[] args) throws Exception {

        if (args.length < 3) {
            System.out.println("Usage: java -jar excel-encryptor.jar <input> <output> <password>");
            return;
        }

        String input = args[0];
        String output = args[1];
        String password = args[2];

        FileInputStream fis = new FileInputStream(input);
        XSSFWorkbook workbook = (XSSFWorkbook) WorkbookFactory.create(fis);

        POIFSFileSystem fs = new POIFSFileSystem();
        EncryptionInfo info = new EncryptionInfo(EncryptionMode.agile);
        Encryptor enc = info.getEncryptor();
        enc.confirmPassword(password);

        OutputStream os = enc.getDataStream(fs);
        workbook.write(os);
        os.close();

        FileOutputStream fos = new FileOutputStream(output);
        fs.writeFilesystem(fos);
        fos.close();

        System.out.println("Encrypted successfully.");
    }
}