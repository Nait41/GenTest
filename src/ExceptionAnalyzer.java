import data.ExceptionList;
import exceptions.ExceptionInfo;
import javafx.beans.value.ChangeListener;
import javafx.beans.value.ObservableValue;
import javafx.event.EventHandler;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.input.MouseEvent;
import javafx.scene.paint.Color;
import javafx.stage.Stage;
import javafx.stage.StageStyle;
import javafx.util.Callback;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class ExceptionAnalyzer extends javafx.application.Application {

    class EditingCell extends TableCell<ExceptionInfo, String> {

        private TextField textField;

        public EditingCell() {
        }

        @Override
        public void startEdit() {
            if (!isEmpty()) {
                super.startEdit();
                createTextField();
                setText(null);
                setGraphic(textField);
                textField.selectAll();
            }
        }

        @Override
        public void cancelEdit() {
            super.cancelEdit();

            setText((String) getItem());
            setGraphic(null);
        }

        @Override
        public void updateItem(String item, boolean empty) {
            super.updateItem(item, empty);

            if (empty) {
                setText(null);
                setGraphic(null);
            } else {
                if (isEditing()) {
                    if (textField != null) {
                        textField.setText(getString());
                    }
                    setText(null);
                    setGraphic(textField);
                } else {
                    setText(getString());
                    setGraphic(null);
                }
            }
        }

        private void createTextField() {
            textField = new TextField(getString());
            textField.setMinWidth(this.getWidth() - this.getGraphicTextGap()* 2);
            textField.focusedProperty().addListener(new ChangeListener<Boolean>(){
                @Override
                public void changed(ObservableValue<? extends Boolean> arg0,
                                    Boolean arg1, Boolean arg2) {
                    if (!arg2) {
                        commitEdit(textField.getText());
                    }
                }
            });
        }

        private String getString() {
            return getItem() == null ? "" : getItem().toString();
        }
    }

    private double xOffset;
    private double yOffset;

    @Override
    public void start(Stage stage) throws IOException {
        FXMLLoader fxmlLoader = new FXMLLoader(Application.class.getResource("panes/exceptionAnalyzer.fxml"));
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
    private TableView<ExceptionInfo> mainTable;

    @FXML
    private Button saveButton;

    @FXML
    void initialize() throws FileNotFoundException {

        TableColumn bacteriaColumn = new TableColumn("Бактерия");
        TableColumn rangeColumn = new TableColumn<>("Средние значение популяции");
        bacteriaColumn.setCellValueFactory(new PropertyValueFactory<ExceptionInfo, String>("bacteria"));
        rangeColumn.setCellValueFactory(new PropertyValueFactory<ExceptionInfo, String>("range"));
        rangeColumn.setPrefWidth(200);
        mainTable.getColumns().add(bacteriaColumn);
        mainTable.getColumns().add(rangeColumn);
        mainTable.getColumns().remove(0);
        mainTable.getColumns().remove(0);
        mainTable.setEditable(true);
        for(int i = 0; i< ExceptionList.exceptBact.size(); i++)
        {
            ExceptionInfo exceptionInfo = new ExceptionInfo();
            exceptionInfo.setBacteria(ExceptionList.exceptBact.get(i).get(0));
            mainTable.getItems().add(exceptionInfo);
        }
        Callback<TableColumn, TableCell> cellFactory =
                new Callback<TableColumn, TableCell>() {
                    public TableCell call(TableColumn p) {
                        return new EditingCell();
                    }
                };
        rangeColumn.setCellFactory(cellFactory);

        rangeColumn.setOnEditCommit(
                new EventHandler<TableColumn.CellEditEvent<ExceptionInfo, String>>() {
                    @Override
                    public void handle(TableColumn.CellEditEvent<ExceptionInfo, String> t) {
                        ExceptionList.exceptBact.get(t.getTablePosition().getRow()).add(1, t.getNewValue());
                    }
                }
        );

        FileInputStream closeStream = new FileInputStream("C:\\Program Files\\gentest_obr\\logout.png");
        Image closeImage = new Image(closeStream);
        ImageView closeView = new ImageView(closeImage);
        closeButton.graphicProperty().setValue(closeView);

        Tooltip closeStart = new Tooltip();
        closeStart.setText("Нажмите, для того, чтобы закрыть окно");
        closeStart.setStyle("-fx-text-fill: turquoise;");
        closeButton.setTooltip(closeStart);

        saveButton.setOnAction(actionEvent -> {
            ErrorController errorController = new ErrorController();
            MainController.errorMessageStr="Для принятия изменений повторно проведите проверку...";
            try {
                errorController.start(new Stage());
            } catch (IOException e) {
                e.printStackTrace();
            }
            new Thread(){
                @Override
                public void run(){
                    File file = new File("C:\\Program Files\\gentest_obr\\algs.xlsx");
                    String filePath = file.getPath();
                    Workbook workbook = null;
                    try {
                        workbook = new XSSFWorkbook(new FileInputStream(filePath));
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                    for(int j = 0; j < ExceptionList.exceptBact.size();j++) {
                        if (ExceptionList.exceptBact.get(j).size() > 1) {
                            workbook.getSheetAt(0).createRow(workbook.getSheetAt(0).getPhysicalNumberOfRows()).createCell(0)
                                    .setCellValue(ExceptionList.exceptBact.get(j).get(0));
                            workbook.getSheetAt(0).getRow(workbook.getSheetAt(0).getPhysicalNumberOfRows()-1).createCell(1)
                                    .setCellValue(ExceptionList.exceptBact.get(j).get(1));
                            workbook.getSheetAt(0).getRow(workbook.getSheetAt(0).getPhysicalNumberOfRows()-1).createCell(2)
                                    .setCellValue("среднее значение");
                        }
                    }
                    try {
                        workbook.write(new FileOutputStream(new File("C:\\Program Files\\gentest_obr\\algs.xlsx")));
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                    try {
                        workbook.close();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
            }.start();
            Stage stage = (Stage) closeButton.getScene().getWindow();
            stage.close();
        });

        closeButton.setOnAction(actionEvent -> {
            Stage stage = (Stage) closeButton.getScene().getWindow();
            stage.close();
        });
    }
}