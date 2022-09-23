import data.AlgsData;
import data.InfoList;
import exceptions.ExceptionInfo;
import javafx.beans.value.ChangeListener;
import javafx.beans.value.ObservableValue;
import javafx.event.ActionEvent;
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
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.concurrent.atomic.AtomicInteger;

public class AlgsTableController extends javafx.application.Application {

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
        FXMLLoader fxmlLoader = new FXMLLoader(Application.class.getResource("panes/algsTable.fxml"));
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
    private TableView<AlgsData> mainTable;

    @FXML
    private Button saveButton;

    @FXML
    private Button addRowButton;

    @FXML
    private Button removeRowButton;

    @FXML
    void initialize() throws IOException, InvalidFormatException {

        TableColumn bacteriaColumn = new TableColumn("Бактерия");
        bacteriaColumn.setCellValueFactory(new PropertyValueFactory<AlgsData, String>("bacteria"));
        mainTable.getColumns().add(bacteriaColumn);
        mainTable.setPrefWidth(200);

        TableColumn rangeColumn = new TableColumn<>("Среднее значение популяции");
        rangeColumn.setCellValueFactory(new PropertyValueFactory<AlgsData, String>("range"));
        mainTable.getColumns().add(rangeColumn);
        rangeColumn.setPrefWidth(200);

        TableColumn rangeInterpretationColumn = new TableColumn<>("Краткая интерпретация");
        rangeInterpretationColumn.setCellValueFactory(new PropertyValueFactory<AlgsData, String>("rangeInterpretation"));
        mainTable.getColumns().add(rangeInterpretationColumn);
        rangeInterpretationColumn.setPrefWidth(200);

        TableColumn genusColumn = new TableColumn<>("Род бактерии");
        genusColumn.setCellValueFactory(new PropertyValueFactory<AlgsData, String>("genus"));
        mainTable.getColumns().add(genusColumn);
        genusColumn.setPrefWidth(200);

        TableColumn descriptionColumn = new TableColumn<>("Описание");
        descriptionColumn.setCellValueFactory(new PropertyValueFactory<AlgsData, String>("description"));
        mainTable.getColumns().add(descriptionColumn);
        descriptionColumn.setPrefWidth(200);


        mainTable.getColumns().remove(0);
        mainTable.getColumns().remove(0);
        mainTable.setEditable(true);
        mainTable.setMinWidth(1020);

        InfoList infoList = new InfoList();
        AlgOpen algOpen = new AlgOpen(infoList);

        int startRow = infoList.algs.size();

        for(int i = 0; i< infoList.algs.size(); i++)
        {
            if (infoList.algs.get(i).get(0) != null)
            {
                AlgsData algsData = new AlgsData();
                algsData.setBacteria(infoList.algs.get(i).get(0));
                algsData.setRange(infoList.algs.get(i).get(1));
                algsData.setRangeInterpretation(infoList.algs.get(i).get(2));
                if (infoList.algs.get(i).size() > 3)
                {
                    algsData.setGenus(infoList.algs.get(i).get(3));
                }
                else
                {
                    infoList.algs.get(i).add("");
                    algsData.setGenus(infoList.algs.get(i).get(3));
                }
                if (infoList.algs.get(i).size() > 4)
                {
                    algsData.setDescription(infoList.algs.get(i).get(4));
                }
                else
                {
                    infoList.algs.get(i).add("");
                    algsData.setDescription(infoList.algs.get(i).get(4));
                }
                mainTable.getItems().add(algsData);
            }
        }

        Callback<TableColumn, TableCell> cellFactoryForBacteria =
                new Callback<TableColumn, TableCell>() {
                    public TableCell call(TableColumn p) {
                        return new EditingCell();
                    }
                };

        Callback<TableColumn, TableCell> cellFactoryForRange =
                new Callback<TableColumn, TableCell>() {
                    public TableCell call(TableColumn p) {
                        return new EditingCell();
                    }
                };

        Callback<TableColumn, TableCell> cellFactoryForRangeInterpretation =
                new Callback<TableColumn, TableCell>() {
                    public TableCell call(TableColumn p) {
                        return new EditingCell();
                    }
                };

        Callback<TableColumn, TableCell> cellFactoryForGenus =
                new Callback<TableColumn, TableCell>() {
                    public TableCell call(TableColumn p) {
                        return new EditingCell();
                    }
                };

        Callback<TableColumn, TableCell> cellFactoryForDescription =
                new Callback<TableColumn, TableCell>() {
                    public TableCell call(TableColumn p) {
                        return new EditingCell();
                    }
                };

        bacteriaColumn.setCellFactory(cellFactoryForBacteria);
        rangeColumn.setCellFactory(cellFactoryForRange);
        rangeInterpretationColumn.setCellFactory(cellFactoryForRangeInterpretation);
        genusColumn.setCellFactory(cellFactoryForGenus);
        descriptionColumn.setCellFactory(cellFactoryForDescription);

        bacteriaColumn.setOnEditCommit(
                new EventHandler<TableColumn.CellEditEvent<ExceptionInfo, String>>() {
                    @Override
                    public void handle(TableColumn.CellEditEvent<ExceptionInfo, String> t) {
                        infoList.algs.get(t.getTablePosition().getRow()).set(0, t.getNewValue());
                    }
                }
        );

        rangeColumn.setOnEditCommit(
                new EventHandler<TableColumn.CellEditEvent<ExceptionInfo, String>>() {
                    @Override
                    public void handle(TableColumn.CellEditEvent<ExceptionInfo, String> t) {
                        infoList.algs.get(t.getTablePosition().getRow()).set(1, t.getNewValue());
                    }
                }
        );

        rangeInterpretationColumn.setOnEditCommit(
                new EventHandler<TableColumn.CellEditEvent<ExceptionInfo, String>>() {
                    @Override
                    public void handle(TableColumn.CellEditEvent<ExceptionInfo, String> t) {
                        infoList.algs.get(t.getTablePosition().getRow()).set(2, t.getNewValue());
                    }
                }
        );

        genusColumn.setOnEditCommit(
                new EventHandler<TableColumn.CellEditEvent<ExceptionInfo, String>>() {
                    @Override
                    public void handle(TableColumn.CellEditEvent<ExceptionInfo, String> t) {
                        if(infoList.algs.get(t.getTablePosition().getRow()).size()>3){
                            infoList.algs.get(t.getTablePosition().getRow()).set(3, t.getNewValue());
                        } else {
                            infoList.algs.get(t.getTablePosition().getRow()).add(3, t.getNewValue());
                        }
                    }
                }
        );

        descriptionColumn.setOnEditCommit(
                new EventHandler<TableColumn.CellEditEvent<ExceptionInfo, String>>() {
                    @Override
                    public void handle(TableColumn.CellEditEvent<ExceptionInfo, String> t) {
                        if(infoList.algs.get(t.getTablePosition().getRow()).size()>4){
                            infoList.algs.get(t.getTablePosition().getRow()).set(4, t.getNewValue());
                        } else {
                            infoList.algs.get(t.getTablePosition().getRow()).add(4, t.getNewValue());
                        }
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

        FileInputStream addStream = new FileInputStream("C:\\Program Files\\gentest_obr\\addAlgs.png");
        Image addImage = new Image(addStream);
        ImageView addView = new ImageView(addImage);
        addRowButton.graphicProperty().setValue(addView);

        Tooltip addTip = new Tooltip();
        addTip.setText("Нажмите, для того, чтобы добавить новую строку");
        addTip.setStyle("-fx-text-fill: turquoise;");
        addRowButton.setTooltip(addTip);

        FileInputStream removeStream = new FileInputStream("C:\\Program Files\\gentest_obr\\removeAlgs.png");
        Image removeImage = new Image(removeStream);
        ImageView removeView = new ImageView(removeImage);
        removeRowButton.graphicProperty().setValue(removeView);

        Tooltip removeTip = new Tooltip();
        removeTip.setText("Нажмите, для того, чтобы удалить выбранную строку");
        removeTip.setStyle("-fx-text-fill: turquoise;");
        removeRowButton.setTooltip(removeTip);

        removeRowButton.setOnAction(ActionEvent -> {
            if (mainTable.getSelectionModel().getSelectedIndex() == -1)
            {
                MainController.errorMessageStr = "Вы не выбрали строку для удаления";
                ErrorController errorController = new ErrorController();
                try {
                    errorController.start(new Stage());
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            else{
                infoList.algs.remove(mainTable.getSelectionModel().getSelectedIndex());
                mainTable.getItems().remove(mainTable.getSelectionModel().getSelectedIndex());
            }
        });

        addRowButton.setOnAction(ActionEvent -> {;
            AlgsData algsData = new AlgsData();
            if (mainTable.getSelectionModel().getSelectedIndex() == -1)
            {
                MainController.errorMessageStr = "Вы не выбрали строку, после которой должна вставиться новая";
                ErrorController errorController = new ErrorController();
                try {
                    errorController.start(new Stage());
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            else{
                infoList.algs.add(mainTable.getSelectionModel().getSelectedIndex()+1, new ArrayList<>());
                infoList.algs.get(mainTable.getSelectionModel().getSelectedIndex()+1).add("");
                infoList.algs.get(mainTable.getSelectionModel().getSelectedIndex()+1).add("");
                infoList.algs.get(mainTable.getSelectionModel().getSelectedIndex()+1).add("");
                infoList.algs.get(mainTable.getSelectionModel().getSelectedIndex()+1).add("");
                infoList.algs.get(mainTable.getSelectionModel().getSelectedIndex()+1).add("");
                mainTable.getItems().add(mainTable.getSelectionModel().getSelectedIndex()+1, algsData);
            }
        });

        saveButton.setOnAction(actionEvent -> {
            File file = new File("C:\\Program Files\\gentest_obr\\algs.xlsx");
            String filePath = file.getPath();
            Workbook workbook = null;
            try {
                workbook = new XSSFWorkbook(new FileInputStream(filePath));
            } catch (IOException e) {
                e.printStackTrace();
            }
            for (int i = 0;i < startRow - infoList.algs.size() + infoList.algs.size();i++)
            {
                workbook.getSheetAt(0).removeRow(workbook.getSheetAt(0).getRow(i));
            }
            for(int j = 0; j < infoList.algs.size();j++) {
                if (infoList.algs.get(j).size() > 0){
                    workbook.getSheetAt(0).createRow(j).createCell(0).setCellValue(infoList.algs.get(j).get(0));
                }
                if (infoList.algs.get(j).size() > 1){
                    workbook.getSheetAt(0).getRow(j).createCell(1).setCellValue(infoList.algs.get(j).get(1));
                }
                if (infoList.algs.get(j).size() > 2){
                    workbook.getSheetAt(0).getRow(j).createCell(2).setCellValue(infoList.algs.get(j).get(2));
                }
                if (infoList.algs.get(j).size() > 3){
                    workbook.getSheetAt(0).getRow(j).createCell(3).setCellValue(infoList.algs.get(j).get(3));
                }
                if (infoList.algs.get(j).size() > 4){
                    workbook.getSheetAt(0).getRow(j).createCell(4).setCellValue(infoList.algs.get(j).get(4));
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
            Stage stage = (Stage) closeButton.getScene().getWindow();
            stage.close();

        });

        closeButton.setOnAction(actionEvent -> {
            Stage stage = (Stage) closeButton.getScene().getWindow();
            stage.close();
        });
    }
}