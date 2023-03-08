package com.rlung;

import com.rlung.util.ExcelUtils;

import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

public class Main {
    public static void main(String[] args) throws IOException, SQLException {

        String jdbcURL = "jdbc:sqlserver://localhost:1433;databaseName=db;trustServerCertificate=true";
        String user = "";
        String pwd = "";
        String tableName = "";
        String inputPath = "inputPath";
        String outputPath = "outputPath";


        tableName = ExcelUtils.getTableNameByReadExcel(inputPath);
        String sql1 = "SELECT TABLE_NAME,COLUMN_NAME,DATA_TYPE FROM Information_Schema.COLUMNS WHERE TABLE_NAME IN (" + tableName + ")";
        String sql2 = "SELECT TABLE_NAME,COLUMN_NAME,DATA_TYPE FROM Information_Schema.COLUMNS";
        Connection conn = null;

        try {

            conn = DriverManager.getConnection(jdbcURL, user, pwd);
            ExcelUtils.writeToExcel(conn, sql1, sql2, "The part of DB document", "All DB document", outputPath);

        } catch (IOException | SQLException e) {
            e.printStackTrace();
        } finally {
            //關閉連線
            if (conn != null) {
                try {
                    conn.close();
                } catch (SQLException e) {
                    e.printStackTrace();
                }
            }
        }
    }
}