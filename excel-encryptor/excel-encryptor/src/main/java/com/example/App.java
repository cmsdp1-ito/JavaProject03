package com.example;

import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.poifs.crypt.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.*;
import java.util.Base64;

public class App {
    public static void main(String[] args) throws Exception {

        if (args.length < 3) {
            // System.out.println("Usage: java -jar excel-encryptor.jar <output> <password> <base64>");
            return;
        }

        String output = args[0];
        String password = args[1];
        String base64 = args[2];

        // ★ C# から渡された Base64 をデコード
        byte[] excelBytes = Base64.getDecoder().decode(base64);

        // ★ メモリ上の Excel を読み込む
        ByteArrayInputStream bais = new ByteArrayInputStream(excelBytes);
        XSSFWorkbook workbook = (XSSFWorkbook) WorkbookFactory.create(bais);

        // ★ 暗号化して保存
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

        // System.out.println("Encrypted Excel created.");
    }
}