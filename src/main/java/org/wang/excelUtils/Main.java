package org.wang.excelUtils;

import com.alibaba.excel.read.metadata.ReadWorkbook;
import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.geometry.Pos;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.layout.Background;
import javafx.scene.layout.Border;
import javafx.scene.layout.GridPane;
import javafx.scene.text.Font;
import javafx.stage.DirectoryChooser;
import javafx.stage.Stage;
import org.wang.excelUtils.service.ExcelUtils;


import java.io.File;
import java.nio.file.Files;
import java.nio.file.Paths;

public class Main extends Application {

    @Override
    public void start(Stage primaryStage) throws Exception{


        TextField textField = new TextField();
        final Button chooserButton = new Button("选择");

        final ToggleGroup group = new ToggleGroup();
        RadioButton radioButton1 = new RadioButton(".xlsx");
        radioButton1.setUserData(".xlsx");
        radioButton1.setSelected(true);
        radioButton1.setToggleGroup(group);
        RadioButton radioButton2 = new RadioButton(".xls");
        radioButton2.setUserData(".xls");
        radioButton2.setToggleGroup(group);

        final Button doActionButton = new Button("合并");

        chooserButton.setOnAction(event -> {
            DirectoryChooser directoryChooser = new DirectoryChooser();
            directoryChooser.setTitle("选择路径");

            File selectedDir =  directoryChooser.showDialog(primaryStage);
            if(selectedDir!=null){
                textField.setText(selectedDir.getAbsolutePath());
            }

        });

        doActionButton.setOnAction(event -> {
            String path = textField.getText();
            if(path==null||path.isEmpty()|| !Files.isDirectory(Paths.get(path))){
                Alert alert = new Alert(Alert.AlertType.INFORMATION);
                alert.setTitle("提示信息");
                alert.setHeaderText(null);
                alert.setContentText("请选择需要合并的文件夹");
                alert.showAndWait();
            }
            String fileType =(String) group.getSelectedToggle().getUserData();
            ExcelUtils.unionWorkBook(path,fileType);

        });





        GridPane gridPane = new GridPane();
        gridPane.setAlignment(Pos.CENTER);
        gridPane.setVgap(4);
        gridPane.setHgap(4);
        gridPane.add(textField,0,0,3,1);
        gridPane.add(chooserButton,3,0,1,1);
        gridPane.add(radioButton1,0,1,1,1);
        gridPane.add(radioButton2,1,1,1,1);
        gridPane.add(doActionButton,0,2);
        primaryStage.setTitle("小脑弟专用excel合并工具");
        primaryStage.setScene(new Scene(gridPane));
        primaryStage.setHeight(300);
        primaryStage.setWidth(400);
        primaryStage.show();
    }


    public static void main(String[] args) {
        launch(args);
    }
}
