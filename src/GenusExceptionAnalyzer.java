import data.ExceptionList;
import exceptions.ExceptionInfo;
import exceptions.GenusExceptionInfo;
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
import java.util.ArrayList;

public class GenusExceptionAnalyzer extends javafx.application.Application {

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
        FXMLLoader fxmlLoader = new FXMLLoader(Application.class.getResource("panes/genusExceptionAnalyzer.fxml"));
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
    private TableView<GenusExceptionInfo> mainTable;

    @FXML
    private Button saveButton;

    public static boolean genusException = false;

    @FXML
    void initialize() throws FileNotFoundException {

        if(MainController.hintsOption){
            Tooltip closeStart = new Tooltip();
            closeStart.setText("Нажмите, для того, чтобы закрыть окно");
            closeStart.setStyle("-fx-text-fill: turquoise;");
            closeButton.setTooltip(closeStart);
        }

        TableColumn bacteriaColumn = new TableColumn("Бактерия");
        TableColumn genusColumn = new TableColumn<>("Род бактерии");
        bacteriaColumn.setCellValueFactory(new PropertyValueFactory<GenusExceptionInfo, String>("bacteria"));
        genusColumn.setCellValueFactory(new PropertyValueFactory<GenusExceptionInfo, String>("genus"));
        genusColumn.setPrefWidth(200);
        mainTable.getColumns().add(bacteriaColumn);
        mainTable.getColumns().add(genusColumn);
        mainTable.getColumns().remove(0);
        mainTable.getColumns().remove(0);
        mainTable.setEditable(true);
        for(int i = 0; i< ExceptionList.genusExceptBact.size(); i++)
        {
            GenusExceptionInfo genusExceptionInfo = new GenusExceptionInfo();
            genusExceptionInfo.setBacteria(ExceptionList.genusExceptBact.get(i).get(0));
            mainTable.getItems().add(genusExceptionInfo);
        }
        Callback<TableColumn, TableCell> cellFactory =
                new Callback<TableColumn, TableCell>() {
                    public TableCell call(TableColumn p) {
                        return new EditingCell();
                    }
                };
        genusColumn.setCellFactory(cellFactory);

        genusColumn.setOnEditCommit(
                new EventHandler<TableColumn.CellEditEvent<ExceptionInfo, String>>() {
                    @Override
                    public void handle(TableColumn.CellEditEvent<ExceptionInfo, String> t) {
                        ExceptionList.genusExceptBact.get(t.getTablePosition().getRow()).add(1, t.getNewValue());
                        mainTable.getItems().get(t.getTablePosition().getRow()).setGenus(t.getNewValue());
                    }
                }
        );

        FileInputStream closeStream = new FileInputStream(Application.rootDirPath + "\\logout.png");
        Image closeImage = new Image(closeStream);
        ImageView closeView = new ImageView(closeImage);
        closeButton.graphicProperty().setValue(closeView);

        saveButton.setOnAction(actionEvent -> {
            new Thread(){
                @Override
                public void run(){
                    File file = new File(Application.rootDirPath + "\\algs.xlsx");
                    String filePath = file.getPath();
                    Workbook workbook = null;
                    try {
                        workbook = new XSSFWorkbook(new FileInputStream(filePath));
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                    ArrayList<String> checkCurrentBacter = new ArrayList<>();
                    ArrayList<ArrayList<String>> missingBacter = new ArrayList<>();
                    int countMissing = 0;
                    for(int j = 0; j < ExceptionList.genusExceptBact.size();j++) {
                        for(int i = 0; i < workbook.getSheetAt(0).getPhysicalNumberOfRows();i++)
                        {
                            if(ExceptionList.genusExceptBact.get(j).size()>2){
                                if (ExceptionList.genusExceptBact.get(j).get(0).equals(workbook.getSheetAt(0).getRow(i).getCell(0).getStringCellValue())) {
                                    workbook.getSheetAt(0).getRow(i).createCell(3).setCellValue(ExceptionList.genusExceptBact.get(j).get(1));
                                    if (!checkCurrentBacter.contains(ExceptionList.genusExceptBact.get(j).get(0))){
                                        checkCurrentBacter.add(ExceptionList.genusExceptBact.get(j).get(0));
                                    }
                                }
                            } else {
                                missingBacter.add(new ArrayList<>());
                                missingBacter.get(countMissing).add(ExceptionList.genusExceptBact.get(j).get(0));
                                countMissing++;
                            }
                        }
                    }
                    if(ExceptionList.genusExceptBact.size() == checkCurrentBacter.size())
                    {
                        MainController.genusOption = false;
                        ExceptionList.genusExceptBact = null;
                    } else
                    {
                        ExceptionList.genusExceptBact = missingBacter;
                    }
                    try {
                        workbook.write(new FileOutputStream(new File(Application.rootDirPath + "\\algs.xlsx")));
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                    try {
                        workbook.close();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }}.start();
            Stage stage = (Stage) closeButton.getScene().getWindow();
            stage.close();
        });

        closeButton.setOnAction(actionEvent -> {
            Stage stage = (Stage) closeButton.getScene().getWindow();
            stage.close();
        });
    }
}