import javafx.animation.KeyFrame;
import javafx.animation.Timeline;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.ToggleButton;
import javafx.scene.control.Tooltip;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.input.MouseEvent;
import javafx.scene.paint.Color;
import javafx.scene.paint.Paint;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import javafx.stage.StageStyle;
import javafx.util.Duration;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import javax.swing.*;
import javax.xml.soap.Text;
import java.io.*;

public class OptionsController extends javafx.application.Application{
    private double xOffset;
    private double yOffset;

    @Override
    public void start(Stage stage) throws IOException {
        FXMLLoader fxmlLoader = new FXMLLoader(Application.class.getResource("panes/options.fxml"));
        Scene scene = new Scene(fxmlLoader.load());
        scene.setFill(Color.TRANSPARENT);
        stage.initStyle(StageStyle.TRANSPARENT);
        scene.setOnMousePressed(new EventHandler<MouseEvent>() {
            @Override
            public void handle(MouseEvent event) {
                xOffset = stage.getX() - event.getScreenX();
                yOffset = stage.getY() - event.getScreenY();
            }
        });
        scene.setOnMouseDragged(new EventHandler<MouseEvent>() {
            @Override
            public void handle(MouseEvent event) {
                stage.setX(event.getScreenX() + xOffset);
                stage.setY(event.getScreenY() + yOffset);
            }
        });
        stage.getIcons().add(new Image("file:///" + Application.rootDirPath + "\\AppIcon.png"));
        stage.setScene(scene);
        stage.show();
    }

    @FXML
    private Button closeButton;

    @FXML
    private Button obrLoadButton;

    @FXML
    private Button obrUnloadButton;

    @FXML
    private Button obr_1_LoadButton;

    @FXML
    private Button obr_1_UnloadButton;

    @FXML
    private Button obr_2_LoadButton;

    @FXML
    private Button obr_2_UnloadButton;

    @FXML
    private Button obr_3_LoadButton;

    @FXML
    private Button obr_3_UnloadButton;

    @FXML
    private Button obr_4_LoadButton;

    @FXML
    private Button obr_4_UnloadButton;

    @FXML
    private ToggleButton descriptionToggle;

    @FXML
    private ToggleButton genusToggle;

    @FXML
    private ToggleButton missingToggle;

    @FXML
    private ToggleButton hintsToggle;

    @FXML
    private ToggleButton mediumRangeToggle;

    @FXML
    private Button unloadAlgsButton;

    @FXML
    private Button loadAlgsButton;

    @FXML
    void initialize() throws FileNotFoundException {

        if(MainController.hintsOption){
            Tooltip obrLoadTip = new Tooltip();
            obrLoadTip.setText("Нажмите, для того, чтобы загрузить расширенный шаблон");
            obrLoadTip.setStyle("-fx-text-fill: #cf6400;");
            obrLoadButton.setTooltip(obrLoadTip);

            Tooltip obr_1LoadTip = new Tooltip();
            obr_1LoadTip.setText("Нажмите, для того, чтобы загрузить шаблон краткой версии ур. микробиома");
            obr_1LoadTip.setStyle("-fx-text-fill: #cf6400;");
            obr_1_LoadButton.setTooltip(obr_1LoadTip);

            Tooltip obr_2LoadTip = new Tooltip();
            obr_2LoadTip.setText("Нажмите, для того, чтобы загрузить шаблон краткой версии микробиома");
            obr_2LoadTip.setStyle("-fx-text-fill: #cf6400;");
            obr_2_LoadButton.setTooltip(obr_2LoadTip);

            Tooltip obr_3LoadTip = new Tooltip();
            obr_3LoadTip.setText("Нажмите, для того, чтобы загрузить шаблон взрослый стандартный");
            obr_3LoadTip.setStyle("-fx-text-fill: #cf6400;");
            obr_3_LoadButton.setTooltip(obr_3LoadTip);

            Tooltip obr_4LoadTip = new Tooltip();
            obr_4LoadTip.setText("Нажмите, для того, чтобы загрузить шаблон микробиома кишечника");
            obr_4LoadTip.setStyle("-fx-text-fill: #cf6400;");
            obr_4_LoadButton.setTooltip(obr_4LoadTip);

            Tooltip obrUnloadTip = new Tooltip();
            obrUnloadTip.setText("Нажмите, для того, чтобы сохранить расширенный шаблон");
            obrUnloadTip.setStyle("-fx-text-fill: #cf6400;");
            obrUnloadButton.setTooltip(obrUnloadTip);

            Tooltip obr_1UnloadTip = new Tooltip();
            obr_1UnloadTip.setText("Нажмите, для того, чтобы сохранить шаблон краткой версии ур. микробиома");
            obr_1UnloadTip.setStyle("-fx-text-fill: #cf6400;");
            obr_1_UnloadButton.setTooltip(obr_1UnloadTip);

            Tooltip obr_2UnloadTip = new Tooltip();
            obr_2UnloadTip.setText("Нажмите, для того, чтобы сохранить шаблон краткой версии микробиома");
            obr_2UnloadTip.setStyle("-fx-text-fill: #cf6400;");
            obr_2_UnloadButton.setTooltip(obr_2UnloadTip);

            Tooltip obr_3UnloadTip = new Tooltip();
            obr_3UnloadTip.setText("Нажмите, для того, чтобы сохранить шаблон взрослый стандартный");
            obr_3UnloadTip.setStyle("-fx-text-fill: #cf6400;");
            obr_3_UnloadButton.setTooltip(obr_3UnloadTip);

            Tooltip obr_4UnloadTip = new Tooltip();
            obr_4UnloadTip.setText("Нажмите, для того, чтобы сохранить шаблон микробиома кишечника");
            obr_4UnloadTip.setStyle("-fx-text-fill: #cf6400;");
            obr_4_UnloadButton.setTooltip(obr_4UnloadTip);

            Tooltip loadAlgsTip = new Tooltip();
            loadAlgsTip.setText("Нажмите, для того, чтобы загрузить новый файл с алгоритмами");
            loadAlgsTip.setStyle("-fx-text-fill: #cf6400;");
            loadAlgsButton.setTooltip(loadAlgsTip);

            Tooltip unloadAlgsTip = new Tooltip();
            unloadAlgsTip.setText("Нажмите, для того, чтобы сохранить текущий файл с алгоритмами");
            unloadAlgsTip.setStyle("-fx-text-fill: #cf6400;");
            unloadAlgsButton.setTooltip(unloadAlgsTip);

            Tooltip closeStart = new Tooltip();
            closeStart.setText("Нажмите, для того, чтобы закрыть окно");
            closeStart.setStyle("-fx-text-fill: #cf6400;");
            closeButton.setTooltip(closeStart);
        }

        FileInputStream obrLoadStream = new FileInputStream(Application.rootDirPath + "\\loadAlgsFile.png");
        Image obrLoadImage = new Image(obrLoadStream);
        ImageView obrLoadView = new ImageView(obrLoadImage);
        obrLoadButton.graphicProperty().setValue(obrLoadView);

        FileInputStream obr_1LoadStream = new FileInputStream(Application.rootDirPath + "\\loadAlgsFile.png");
        Image obr_1LoadImage = new Image(obr_1LoadStream);
        ImageView obr_1LoadView = new ImageView(obr_1LoadImage);
        obr_1_LoadButton.graphicProperty().setValue(obr_1LoadView);

        FileInputStream obr_2LoadStream = new FileInputStream(Application.rootDirPath + "\\loadAlgsFile.png");
        Image obr_2LoadImage = new Image(obr_2LoadStream);
        ImageView obr_2LoadView = new ImageView(obr_2LoadImage);
        obr_2_LoadButton.graphicProperty().setValue(obr_2LoadView);

        FileInputStream obr_3LoadStream = new FileInputStream(Application.rootDirPath + "\\loadAlgsFile.png");
        Image obr_3LoadImage = new Image(obr_3LoadStream);
        ImageView obr_3LoadView = new ImageView(obr_3LoadImage);
        obr_3_LoadButton.graphicProperty().setValue(obr_3LoadView);

        FileInputStream obr_4LoadStream = new FileInputStream(Application.rootDirPath + "\\loadAlgsFile.png");
        Image obr_4LoadImage = new Image(obr_4LoadStream);
        ImageView obr_4LoadView = new ImageView(obr_4LoadImage);
        obr_4_LoadButton.graphicProperty().setValue(obr_4LoadView);

        FileInputStream obrUnloadStream = new FileInputStream(Application.rootDirPath + "\\saveAlgsFile.png");
        Image obrUnloadImage = new Image(obrUnloadStream);
        ImageView obrUnloadView = new ImageView(obrUnloadImage);
        obrUnloadButton.graphicProperty().setValue(obrUnloadView);

        FileInputStream obr_1UnloadStream = new FileInputStream(Application.rootDirPath + "\\saveAlgsFile.png");
        Image obr_1UnloadImage = new Image(obr_1UnloadStream);
        ImageView obr_1UnloadView = new ImageView(obr_1UnloadImage);
        obr_1_UnloadButton.graphicProperty().setValue(obr_1UnloadView);

        FileInputStream obr_2UnloadStream = new FileInputStream(Application.rootDirPath + "\\saveAlgsFile.png");
        Image obr_2UnloadImage = new Image(obr_2UnloadStream);
        ImageView obr_2UnloadView = new ImageView(obr_2UnloadImage);
        obr_2_UnloadButton.graphicProperty().setValue(obr_2UnloadView);

        FileInputStream obr_3UnloadStream = new FileInputStream(Application.rootDirPath + "\\saveAlgsFile.png");
        Image obr_3UnloadImage = new Image(obr_3UnloadStream);
        ImageView obr_3UnloadView = new ImageView(obr_3UnloadImage);
        obr_3_UnloadButton.graphicProperty().setValue(obr_3UnloadView);

        FileInputStream obr_4UnloadStream = new FileInputStream(Application.rootDirPath + "\\saveAlgsFile.png");
        Image obr_4UnloadImage = new Image(obr_4UnloadStream);
        ImageView obr_4UnloadView = new ImageView(obr_4UnloadImage);
        obr_4_UnloadButton.graphicProperty().setValue(obr_4UnloadView);

        FileInputStream loadAlgsStream = new FileInputStream(Application.rootDirPath + "\\loadAlgsFile.png");
        Image loadAlgsImage = new Image(loadAlgsStream);
        ImageView loadAlgsView = new ImageView(loadAlgsImage);
        loadAlgsButton.graphicProperty().setValue(loadAlgsView);

        FileInputStream unloadAlgsStream = new FileInputStream(Application.rootDirPath + "\\saveAlgsFile.png");
        Image unloadAlgsImage = new Image(unloadAlgsStream);
        ImageView unloadAlgsView = new ImageView(unloadAlgsImage);
        unloadAlgsButton.graphicProperty().setValue(unloadAlgsView);

        FileInputStream closeStream = new FileInputStream(Application.rootDirPath + "\\logout.png");
        Image closeImage = new Image(closeStream);
        ImageView closeView = new ImageView(closeImage);
        closeButton.graphicProperty().setValue(closeView);

        closeButton.setOnAction(actionEvent -> {
            Stage stage = (Stage) closeButton.getScene().getWindow();
            stage.close();
        });

        loadAlgsButton.setOnAction(ActionEvent -> {
            FileChooser fileChooser = new FileChooser();
            File file = fileChooser.showOpenDialog(new Stage());
            try {
                XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(file));
                workbook.write(new FileOutputStream(new File(Application.rootDirPath + "\\algs.xlsx")));
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        });

        obrLoadButton.setOnAction(ActionEvent -> {
            FileChooser fileChooser = new FileChooser();
            File file = fileChooser.showOpenDialog(new Stage());
            try {
                XWPFDocument workbook = new XWPFDocument(new FileInputStream(file));
                workbook.write(new FileOutputStream(new File(Application.rootDirPath + "\\obr.xlsx")));
                workbook.write(new FileOutputStream(new File(Application.rootDirPath + "\\exceptionCheckObrFile\\obr.xlsx")));
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        });

        obr_1_LoadButton.setOnAction(ActionEvent -> {
            FileChooser fileChooser = new FileChooser();
            File file = fileChooser.showOpenDialog(new Stage());
            try {
                XWPFDocument workbook = new XWPFDocument(new FileInputStream(file));
                workbook.write(new FileOutputStream(new File(Application.rootDirPath + "\\obr_1.xlsx")));
                workbook.write(new FileOutputStream(new File(Application.rootDirPath + "\\exceptionCheckObrFile\\obr_1.xlsx")));
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        });

        obr_2_LoadButton.setOnAction(ActionEvent -> {
            FileChooser fileChooser = new FileChooser();
            File file = fileChooser.showOpenDialog(new Stage());
            try {
                XWPFDocument workbook = new XWPFDocument(new FileInputStream(file));
                workbook.write(new FileOutputStream(new File(Application.rootDirPath + "\\obr_2.xlsx")));
                workbook.write(new FileOutputStream(new File(Application.rootDirPath + "\\exceptionCheckObrFile\\obr_2.xlsx")));
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        });

        obr_3_LoadButton.setOnAction(ActionEvent -> {
            FileChooser fileChooser = new FileChooser();
            File file = fileChooser.showOpenDialog(new Stage());
            try {
                XWPFDocument workbook = new XWPFDocument(new FileInputStream(file));
                workbook.write(new FileOutputStream(new File(Application.rootDirPath + "\\obr_3.xlsx")));
                workbook.write(new FileOutputStream(new File(Application.rootDirPath + "\\exceptionCheckObrFile\\obr_3.xlsx")));
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        });

        obr_4_LoadButton.setOnAction(ActionEvent -> {
            FileChooser fileChooser = new FileChooser();
            File file = fileChooser.showOpenDialog(new Stage());
            try {
                XWPFDocument workbook = new XWPFDocument(new FileInputStream(file));
                workbook.write(new FileOutputStream(new File(Application.rootDirPath + "\\obr_4.xlsx")));
                workbook.write(new FileOutputStream(new File(Application.rootDirPath + "\\exceptionCheckObrFile\\obr_4.xlsx")));
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        });

        unloadAlgsButton.setOnAction(ActionEvent -> {
            DirectoryChooser directoryChooser = new DirectoryChooser();
            File file = directoryChooser.showDialog(new Stage());
            try {
                XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(Application.rootDirPath + "\\algs.xlsx"));
                workbook.write(new FileOutputStream(file + "\\Алгоритмы_метагеном.xlsx"));
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        });

        unloadAlgsButton.setOnAction(ActionEvent -> {
            DirectoryChooser directoryChooser = new DirectoryChooser();
            File file = directoryChooser.showDialog(new Stage());
            try {
                XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(Application.rootDirPath + "\\algs.xlsx"));
                workbook.write(new FileOutputStream(file + "\\Алгоритмы_метагеном.xlsx"));
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        });

        obrUnloadButton.setOnAction(ActionEvent -> {
            DirectoryChooser directoryChooser = new DirectoryChooser();
            File file = directoryChooser.showDialog(new Stage());
            try {
                XWPFDocument workbook = new XWPFDocument(new FileInputStream(Application.rootDirPath + "\\obr.docx"));
                workbook.write(new FileOutputStream(file + "\\Расширенный_шаблон.docx"));
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        });

        obr_1_UnloadButton.setOnAction(ActionEvent -> {
            DirectoryChooser directoryChooser = new DirectoryChooser();
            File file = directoryChooser.showDialog(new Stage());
            try {
                XWPFDocument workbook = new XWPFDocument(new FileInputStream(Application.rootDirPath + "\\obr_1.docx"));
                workbook.write(new FileOutputStream(file + "\\Шаблон_краткой_версии_ур._микробиома.docx"));
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        });

        obr_2_UnloadButton.setOnAction(ActionEvent -> {
            DirectoryChooser directoryChooser = new DirectoryChooser();
            File file = directoryChooser.showDialog(new Stage());
            try {
                XWPFDocument workbook = new XWPFDocument(new FileInputStream(Application.rootDirPath + "\\obr_2.docx"));
                workbook.write(new FileOutputStream(file + "\\Шаблон_краткой_версии_микробиома.docx"));
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        });

        obr_3_UnloadButton.setOnAction(ActionEvent -> {
            DirectoryChooser directoryChooser = new DirectoryChooser();
            File file = directoryChooser.showDialog(new Stage());
            try {
                XWPFDocument workbook = new XWPFDocument(new FileInputStream(Application.rootDirPath + "\\obr_3.docx"));
                workbook.write(new FileOutputStream(file + "\\Шаблон_взрослый_стандартный.docx"));
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        });

        obr_4_UnloadButton.setOnAction(ActionEvent -> {
            DirectoryChooser directoryChooser = new DirectoryChooser();
            File file = directoryChooser.showDialog(new Stage());
            try {
                XWPFDocument workbook = new XWPFDocument(new FileInputStream(Application.rootDirPath + "\\obr_4.docx"));
                workbook.write(new FileOutputStream(file + "\\Шаблон_микробиома_кишечника.docx"));
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        });

        if (MainController.hintsOption){
            hintsToggle.setStyle("-fx-background-color: #cf6400");
            hintsToggle.setTextFill(Paint.valueOf("#ebebeb"));
            hintsToggle.setText("Активно");
        } else
        {
            hintsToggle.setStyle("-fx-background-color: #ebebeb");
            hintsToggle.setTextFill(Paint.valueOf("#cf6400"));
            hintsToggle.setText("Не активно");
        }

        if (MainController.missingOption){
            missingToggle.setStyle("-fx-background-color: #cf6400");
            missingToggle.setTextFill(Paint.valueOf("#ebebeb"));
            missingToggle.setText("Активно");
        } else
        {
            missingToggle.setStyle("-fx-background-color: #ebebeb");
            missingToggle.setTextFill(Paint.valueOf("#cf6400"));
            missingToggle.setText("Не активно");
        }

        if (MainController.mediumRangeOption){
            mediumRangeToggle.setStyle("-fx-background-color: #cf6400");
            mediumRangeToggle.setTextFill(Paint.valueOf("#ebebeb"));
            mediumRangeToggle.setText("Активно");
        } else
        {
            mediumRangeToggle.setStyle("-fx-background-color: #ebebeb");
            mediumRangeToggle.setTextFill(Paint.valueOf("#cf6400"));
            mediumRangeToggle.setText("Не активно");
        }

        if (MainController.genusOption){
            genusToggle.setStyle("-fx-background-color: #cf6400");
            genusToggle.setTextFill(Paint.valueOf("#ebebeb"));
            genusToggle.setText("Активно");;
        } else
        {
            genusToggle.setStyle("-fx-background-color: #ebebeb");
            genusToggle.setTextFill(Paint.valueOf("#cf6400"));
            genusToggle.setText("Не активно");
        }

        if (MainController.descriptionOption){
            descriptionToggle.setStyle("-fx-background-color: #cf6400");
            descriptionToggle.setTextFill(Paint.valueOf("#ebebeb"));
            descriptionToggle.setText("Активно");
        } else
        {
            descriptionToggle.setStyle("-fx-background-color: #ebebeb");
            descriptionToggle.setTextFill(Paint.valueOf("#cf6400"));
            descriptionToggle.setText("Не активно");
        }

        hintsToggle.setOnAction(ActionEvent -> {
            if (hintsToggle.isSelected()){
                hintsToggle.setStyle("-fx-background-color: #cf6400");
                hintsToggle.setTextFill(Paint.valueOf("#ebebeb"));
                hintsToggle.setText("Активно");
                MainController.hintsOption = true;
            } else
            {;
                hintsToggle.setStyle("-fx-background-color: #ebebeb");
                hintsToggle.setTextFill(Paint.valueOf("#cf6400"));
                hintsToggle.setText("Не активно");
                MainController.hintsOption = false;
            }
        });

        missingToggle.setOnAction(ActionEvent -> {
            if (missingToggle.isSelected()){
                missingToggle.setStyle("-fx-background-color: #cf6400");
                missingToggle.setTextFill(Paint.valueOf("#ebebeb"));
                missingToggle.setText("Активно");
                MainController.missingOption = true;
            } else
            {;
                missingToggle.setStyle("-fx-background-color: #ebebeb");
                missingToggle.setTextFill(Paint.valueOf("#cf6400"));
                missingToggle.setText("Не активно");
                MainController.missingOption = false;
            }
        });

        mediumRangeToggle.setOnAction(ActionEvent -> {
            if (mediumRangeToggle.isSelected()){
                mediumRangeToggle.setStyle("-fx-background-color: #cf6400");
                mediumRangeToggle.setTextFill(Paint.valueOf("#ebebeb"));
                mediumRangeToggle.setText("Активно");
                MainController.mediumRangeOption = true;
            } else
            {
                mediumRangeToggle.setStyle("-fx-background-color: #ebebeb");
                mediumRangeToggle.setTextFill(Paint.valueOf("#cf6400"));
                mediumRangeToggle.setText("Не активно");
                MainController.mediumRangeOption = false;
            }
        });

        genusToggle.setOnAction(ActionEvent -> {
            if (genusToggle.isSelected()){
                genusToggle.setStyle("-fx-background-color: #cf6400");
                genusToggle.setTextFill(Paint.valueOf("#ebebeb"));
                genusToggle.setText("Активно");
                MainController.genusOption = true;
            } else
            {
                genusToggle.setStyle("-fx-background-color: #ebebeb");
                genusToggle.setTextFill(Paint.valueOf("#cf6400"));
                genusToggle.setText("Не активно");
                MainController.genusOption = false;
            }
        });

        descriptionToggle.setOnAction(ActionEvent -> {
            if (descriptionToggle.isSelected()){
                descriptionToggle.setStyle("-fx-background-color: #cf6400");
                descriptionToggle.setTextFill(Paint.valueOf("#ebebeb"));
                descriptionToggle.setText("Активно");
                MainController.descriptionOption = true;
            } else
            {
                descriptionToggle.setStyle("-fx-background-color: #ebebeb");
                descriptionToggle.setTextFill(Paint.valueOf("#cf6400"));
                descriptionToggle.setText("Не активно");
                MainController.descriptionOption = false;
            }
        });
    }
}
