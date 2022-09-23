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
import javafx.stage.Stage;
import javafx.stage.StageStyle;

import javax.swing.*;
import javax.xml.soap.Text;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

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
        stage.getIcons().add(new Image("file:///C:\\Program Files\\gentest_obr\\AppIcon.png"));
        stage.setScene(scene);
        stage.show();
    }

    @FXML
    private Button closeButton;

    @FXML
    private ToggleButton descriptionToggle;

    @FXML
    private ToggleButton genusToggle;

    @FXML
    private ToggleButton mediumRangeToggle;

    @FXML
    private Button unloadAlgsButton;

    @FXML
    private Button loadAlgsButton;

    @FXML
    void initialize() throws FileNotFoundException {

        FileInputStream loadAlgsStream = new FileInputStream("C:\\Program Files\\gentest_obr\\loadAlgsFile.png");
        Image loadAlgsImage = new Image(loadAlgsStream);
        ImageView loadAlgsView = new ImageView(loadAlgsImage);
        loadAlgsButton.graphicProperty().setValue(loadAlgsView);

        Tooltip loadAlgsTip = new Tooltip();
        loadAlgsTip.setText("Нажмите, для того, чтобы загрузить новый файл с алгоритмами");
        loadAlgsTip.setStyle("-fx-text-fill: #cf6400;");
        loadAlgsButton.setTooltip(loadAlgsTip);

        FileInputStream unloadAlgsStream = new FileInputStream("C:\\Program Files\\gentest_obr\\saveAlgsFile.png");
        Image unloadAlgsImage = new Image(unloadAlgsStream);
        ImageView unloadAlgsView = new ImageView(unloadAlgsImage);
        unloadAlgsButton.graphicProperty().setValue(unloadAlgsView);

        Tooltip unloadAlgsTip = new Tooltip();
        unloadAlgsTip.setText("Нажмите, для того, чтобы сохранить текущий файл с алгоритмами");
        unloadAlgsTip.setStyle("-fx-text-fill: #cf6400;");
        unloadAlgsButton.setTooltip(unloadAlgsTip);

        FileInputStream closeStream = new FileInputStream("C:\\Program Files\\gentest_obr\\logout.png");
        Image closeImage = new Image(closeStream);
        ImageView closeView = new ImageView(closeImage);
        closeButton.graphicProperty().setValue(closeView);

        Tooltip closeStart = new Tooltip();
        closeStart.setText("Нажмите, для того, чтобы закрыть окно");
        closeStart.setStyle("-fx-text-fill: #cf6400;");
        closeButton.setTooltip(closeStart);

        closeButton.setOnAction(actionEvent -> {
            Stage stage = (Stage) closeButton.getScene().getWindow();
            stage.close();
        });

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
