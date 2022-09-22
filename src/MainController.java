import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.net.URL;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.ResourceBundle;

import data.InfoList;
import fileView.XLXSOpen;
import javafx.animation.KeyFrame;
import javafx.animation.Timeline;
import javafx.beans.value.ChangeListener;
import javafx.beans.value.ObservableValue;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.fxml.FXML;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.ListView;
import javafx.scene.control.Tooltip;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.layout.AnchorPane;
import javafx.scene.shape.Circle;
import javafx.scene.text.Text;
import javafx.stage.DirectoryChooser;
import javafx.stage.Stage;
import javafx.util.Duration;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import static java.lang.Thread.sleep;

public class MainController {
    public static boolean descriptionOption = false;
    public static boolean genusOption = false;
    public static boolean mediumRangeOption = false;
    public static boolean exceptCheck = false;
    public InfoList infoList;
    AlgOpen alg;
    ArrayList<String> content_list = new ArrayList<>();
    List<File> samplePath;
    String selectedSample = "";
    String selectedException = "";
    MainLoader docLoad;
    XLXSOpen xlxsOpen;
    File saveSampleDir;
    boolean checkLoad, checkUnload, checkStart = false;
    int counter, counter_files;
    public static String errorMessageStr = "";

    @FXML
    private ResourceBundle resources;

    @FXML
    private URL location;

    @FXML
    private Button dirLoadButton;

    @FXML
    private Button algsTable;

    @FXML
    private Button dirUnloadButton;

    ArrayList<String> langs = new ArrayList<>();

    @FXML
    private ListView<String> listSample;

    @FXML
    private Text loadStatus;

    @FXML
    private Text loadStatus_end;

    @FXML
    private Text loadStatusFileNumber;

    @FXML
    private Button startButton;

    @FXML
    private AnchorPane mainPanel;

    @FXML
    private Button closeButton;

    @FXML
    private Button exceptionButton;

    @FXML
    private AnchorPane exceptPane;

    @FXML
    private ListView<String> exceptView;

    @FXML
    private Button options;

    public MainController() throws IOException, InvalidFormatException {
    }

    void feelLangs(){
        langs.add("Рассширенный образец");
        langs.add("Краткая версия урогенитального микробиома");
        langs.add("Краткая версия микробиома");
        langs.add("Образец взрослый стандартный");
        langs.add("Микробиом кишечника");
    }

    int getCounter(int rowCount, int currentNumber) {
        Double temp = new Double(100/rowCount);
        return temp.intValue() + currentNumber;
    }

    void feelExceptLangs(){
        if (!exceptView.getItems().contains("Не для всех бактерий определены среднии значения популяции"))
        {
            if(exceptCheck && mediumRangeOption){
                exceptView.getItems().add("Не для всех бактерий определены среднии значения популяции");
            }
        }
        if(!exceptView.getItems().contains("Не для всех бактерий определен род")){
            if(GenusExceptionAnalyzer.genusException && genusOption){
                exceptView.getItems().add("Не для всех бактерий определен род");
            }
        }
        if(!exceptView.getItems().contains("Не все бактерии описаны")){
            if(DescriptionExceptionAnalyzer.descriptionExcept && descriptionOption){
                exceptView.getItems().add("Не все бактерии описаны");
            }
        }
    }

    @FXML
    void initialize() throws FileNotFoundException {

        exceptPane.setVisible(false);

        exceptionButton.setVisible(false);

        FileInputStream optionsStream = new FileInputStream("C:\\Program Files\\gentest_obr\\options.png");
        Image optionsImage = new Image(optionsStream);
        ImageView optionsView = new ImageView(optionsImage);
        options.graphicProperty().setValue(optionsView);

        FileInputStream loadStream = new FileInputStream("C:\\Program Files\\gentest_obr\\load.png");
        Image loadImage = new Image(loadStream);
        ImageView loadView = new ImageView(loadImage);
        dirLoadButton.graphicProperty().setValue(loadView);

        FileInputStream unloadStream = new FileInputStream("C:\\Program Files\\gentest_obr\\unload.png");
        Image unloadImage = new Image(unloadStream);
        ImageView unloadView = new ImageView(unloadImage);
        dirUnloadButton.graphicProperty().setValue(unloadView);

        FileInputStream startStream = new FileInputStream("C:\\Program Files\\gentest_obr\\start.png");
        Image startImage = new Image(startStream);
        ImageView startView = new ImageView(startImage);
        startButton.graphicProperty().setValue(startView);

        FileInputStream closeStream = new FileInputStream("C:\\Program Files\\gentest_obr\\logout.png");
        Image closeImage = new Image(closeStream);
        ImageView closeView = new ImageView(closeImage);
        closeButton.graphicProperty().setValue(closeView);

        FileInputStream exceptionStream = new FileInputStream("C:\\Program Files\\gentest_obr\\exception.png");
        Image exceptionImage = new Image(exceptionStream);
        ImageView exceptionView = new ImageView(exceptionImage);
        exceptionButton.graphicProperty().setValue(exceptionView);

        FileInputStream algsTableStream = new FileInputStream("C:\\Program Files\\gentest_obr\\algsTable.png");
        Image algsTableImage = new Image(algsTableStream);
        ImageView algsTableView = new ImageView(algsTableImage);
        algsTable.graphicProperty().setValue(algsTableView);

        Tooltip tipAlgsTable = new Tooltip();
        tipAlgsTable.setText("Нажмите, для того, чтобы перейти к редактированию таблицы алгоритмов");
        tipAlgsTable.setStyle("-fx-text-fill: turquoise;");
        algsTable.setTooltip(tipAlgsTable);

        Tooltip tipLoad = new Tooltip();
        tipLoad.setText("Выберите папку, в которой находятся xlsx файлы");
        tipLoad.setStyle("-fx-text-fill: turquoise;");
        dirLoadButton.setTooltip(tipLoad);

        Tooltip tipOptions = new Tooltip();
        tipOptions.setText("Нажмите, для того, чтобы перейти в опции");
        tipOptions.setStyle("-fx-text-fill: turquoise;");
        options.setTooltip(tipOptions);

        Tooltip tipUnLoad = new Tooltip();
        tipUnLoad.setText("Выберите папку, в которую должны сохраняться готовые отчеты");
        tipUnLoad.setStyle("-fx-text-fill: turquoise;");
        dirUnloadButton.setTooltip(tipUnLoad);

        Tooltip tipStart = new Tooltip();
        tipStart.setText("Нажмите, для того, чтобы получить готовые отчеты");
        tipStart.setStyle("-fx-text-fill: turquoise;");
        startButton.setTooltip(tipStart);

        Tooltip closeStart = new Tooltip();
        closeStart.setText("Нажмите, для того, чтобы закрыть приложение");
        closeStart.setStyle("-fx-text-fill: turquoise;");
        closeButton.setTooltip(closeStart);

        Tooltip exceptionTip = new Tooltip();
        exceptionTip.setText("Нажмите на кнопку, чтобы посмотреть список проблем");
        exceptionTip.setStyle("-fx-text-fill: turquoise;");
        exceptionButton.setTooltip(exceptionTip);

        exceptView.getSelectionModel().selectedItemProperty().addListener(new ChangeListener<String>() {
            ExceptionAnalyzer exceptionAnalyzer = new ExceptionAnalyzer();
            GenusExceptionAnalyzer genusExceptionAnalyzer = new GenusExceptionAnalyzer();
            DescriptionExceptionAnalyzer descriptionExceptionAnalyzer = new DescriptionExceptionAnalyzer();
            @Override
            public void changed(ObservableValue<? extends String> observable, String oldValue, String newValue) {
                selectedException = exceptView.getSelectionModel().getSelectedItem();
                if(selectedException.equals("Не для всех бактерий определены среднии значения популяции")){
                    try {
                        exceptionAnalyzer.start(new Stage());
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
                if(selectedException.equals("Не для всех бактерий определен род")){
                    try {
                        genusExceptionAnalyzer.start(new Stage());
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
                if(selectedException.equals("Не все бактерии описаны")){
                    try {
                        descriptionExceptionAnalyzer.start(new Stage());
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
            }
        });

        int r = 60;
        startButton.setShape(new Circle(r));
        startButton.setMinSize(r*2, r*2);
        startButton.setMaxSize(r*2, r*2);

        checkLoad = false;
        checkUnload = false;
        feelLangs();
        listSample.getItems().addAll(langs);
        listSample.getSelectionModel().selectedItemProperty().addListener(new ChangeListener<String>() {
            @Override
            public void changed(ObservableValue<? extends String> observableValue, String s, String t1) {
                selectedSample = listSample.getSelectionModel().getSelectedItem();
            }
        });

        closeButton.setOnAction(actionEvent -> {
            Stage stage = (Stage) closeButton.getScene().getWindow();
            stage.close();
        });

        exceptionButton.setOnAction(actionEvent -> {
            if(exceptPane.isVisible()){
                exceptPane.setVisible(false);
            }
            else{
                exceptPane.setVisible(true);
                feelExceptLangs();
                if (exceptView.getItems().size()<2)
                {
                    ExceptionAnalyzer exceptionAnalyzer = new ExceptionAnalyzer();
                    GenusExceptionAnalyzer genusExceptionAnalyzer = new GenusExceptionAnalyzer();
                    if(selectedException.equals("Не для всех бактерий определены среднии значения популяции")){
                        try {
                            exceptionAnalyzer.start(new Stage());
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    }
                    if(selectedException.equals("Не для всех бактерий определен род")){
                        try {
                            genusExceptionAnalyzer.start(new Stage());
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    }
                }
            }
        });

        dirLoadButton.setOnAction(actionEvent -> {
            if(!checkStart)
            {
                loadStatus.setText("");
                loadStatus_end.setText("");
                loadStatusFileNumber.setText("");
                DirectoryChooser directoryChooser = new DirectoryChooser();
                File dir = directoryChooser.showDialog(new Stage());
                File[] file = dir.listFiles();
                samplePath = Arrays.asList(file);
                checkLoad = true;
            }
            else
            {
                errorMessageStr = "Происходит обработка файлов. Повторите попытку попытку позже...";
                ErrorController errorController = new ErrorController();
                try {
                    errorController.start(new Stage());
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        });

        options.setOnAction(ActionEvent -> {
            OptionsController optionsController = new OptionsController();
            try {
                optionsController.start(new Stage());
            } catch (IOException e) {
                e.printStackTrace();
            }
        });

        dirUnloadButton.setOnAction(actionEvent -> {
                    if(!checkStart)
                    {
                        loadStatus.setText("");
                        loadStatus_end.setText("");
                        loadStatusFileNumber.setText("");
                        DirectoryChooser directoryChooser = new DirectoryChooser();
                        saveSampleDir = directoryChooser.showDialog(new Stage());
                        checkUnload = true;

                    }
                    else
                    {
                        errorMessageStr = "Происходит обработка файлов. Повторите попытку попытку позже...";
                        ErrorController errorController = new ErrorController();
                        try {
                            errorController.start(new Stage());
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    }
                }
        );
        startButton.setOnAction(actionEvent -> {
                    if(!checkStart){
                        loadStatus.setText("");
                        loadStatus_end.setText("");
                        loadStatusFileNumber.setText("");
                        if(checkLoad & checkUnload){
                            if(!selectedSample.equals(""))
                            {
                                if(samplePath.size() != 0)
                                {
                                    checkStart = true;
                                    if(selectedSample.equals("Рассширенный образец")){
                                        new Thread(){
                                            @Override
                                            public void run(){
                                                counter_files = 0;
                                                for (int i = 0; i<samplePath.size();i++)
                                                {
                                                    if(samplePath.get(i).getPath().contains(".xlsx"))
                                                    {
                                                        loadStatusFileNumber.setText("Обработка " + i + " файла");
                                                        counter = 0;
                                                        infoList = new InfoList();
                                                        try {
                                                            xlxsOpen = new XLXSOpen(samplePath.get(i));
                                                            docLoad = new MainLoader("obr");
                                                            alg = new AlgOpen(infoList);
                                                        } catch (IOException e) {
                                                            e.printStackTrace();
                                                        } catch (InvalidFormatException e) {
                                                            e.printStackTrace();
                                                        }
                                                        try {
                                                            xlxsOpen.getPhylum(infoList);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                            xlxsOpen.getGenus(infoList);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                            xlxsOpen.getFileName(infoList);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                            xlxsOpen.getSpecies(infoList);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                        } catch (IOException e) {
                                                            e.printStackTrace();
                                                        }
                                                        docLoad.setFileNameForFirst(infoList);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                        docLoad.setPhylum(infoList);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                        docLoad.setRatio(infoList);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                        docLoad.setFiveFormat(infoList, 4);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                        docLoad.setFiveFormat(infoList, 5);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                        docLoad.setFiveFormat(infoList, 6);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                        docLoad.setFourFormat(infoList, 7);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                        docLoad.setThreeDoubleFormat(infoList, 8);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                        docLoad.setThreeDoubleFormat(infoList, 9);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                        docLoad.setThreeDoubleFormat(infoList, 10);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                        docLoad.setFourFormat(infoList,11);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                        docLoad.setAddition(infoList);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                        docLoad.setTwoFormat(infoList, 14);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                        docLoad.setTwoFormat(infoList, 15);
                                                        loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                        try {
                                                            //docLoad.saveSortedTable(infoList, 14, "First");
                                                            //docLoad.saveSortedTable(infoList, 15, "Second");
                                                            docLoad.setTwoFormatWithSer(infoList, 14, SerFileName.FIRST.toString());
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                            docLoad.setTwoFormatWithSer(infoList, 15, SerFileName.SECOND.toString());
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(21, counter)) + " %");
                                                            docLoad.saveFile(infoList, saveSampleDir);
                                                            loadStatus.setText("Загрузка: 100 % ");
                                                        } catch (IOException e) {
                                                            e.printStackTrace();
                                                        } catch (ClassNotFoundException e) {
                                                            e.printStackTrace();
                                                        }
                                                        try {
                                                            docLoad.getClose();
                                                            xlxsOpen.getClose();
                                                        } catch (IOException e) {
                                                            e.printStackTrace();
                                                        }
                                                        if(mediumRangeOption){
                                                            try {
                                                                docLoad.saveObrFile();
                                                            } catch (IOException e) {
                                                                e.printStackTrace();
                                                            }
                                                        }
                                                        counter_files++;
                                                    }
                                                }
                                                loadStatusFileNumber.setText("");
                                                loadStatus_end.setText("Успешно обработано " + counter_files + " файла(ов)!");
                                                checkStart = false;
                                                if(mediumRangeOption || genusOption || descriptionOption){
                                                    exceptionButton.setVisible(true);
                                                }
                                            }
                                        }.start();
                                    } else if(selectedSample.equals("Краткая версия урогенитального микробиома")){
                                        new Thread(){
                                            @Override
                                            public void run(){
                                                counter_files = 0;
                                                for (int i = 0; i<samplePath.size();i++) {
                                                    if(samplePath.get(i).getPath().contains(".xlsx"))
                                                    {
                                                        loadStatusFileNumber.setText("Обработка " + i + " файла");
                                                        counter = 0;
                                                        infoList = new InfoList();
                                                        try {
                                                            xlxsOpen = new XLXSOpen(samplePath.get(i));
                                                            docLoad = new MainLoader("obr_1");
                                                            alg = new AlgOpen(infoList);
                                                        } catch (IOException e) {
                                                            e.printStackTrace();
                                                        } catch (InvalidFormatException e) {
                                                            e.printStackTrace();
                                                        }
                                                        try {
                                                            xlxsOpen.getPhylum(infoList);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(10, counter)) + " %");
                                                            xlxsOpen.getGenus(infoList);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(10, counter)) + " %");
                                                            xlxsOpen.getFileName(infoList);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(10, counter)) + " %");
                                                            xlxsOpen.getSpecies(infoList);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(10, counter)) + " %");
                                                            docLoad.setFileNameForSecond(infoList);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(10, counter)) + " %");
                                                            docLoad.setFourTableFormatForSecond(infoList, 0);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(10, counter)) + " %");
                                                            docLoad.setFourTableFormatForSecond(infoList, 1);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(10, counter)) + " %");
                                                            //docLoad.saveSortedTable(infoList, 1, "Third");
                                                            //docLoad.saveSortedTable(infoList, 2, "Fourth");
                                                            docLoad.setTwoFormatWithSer(infoList, 1, SerFileName.THIRD.toString());
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(10, counter)) + " %");
                                                            docLoad.setTwoFormatWithSer(infoList, 2, SerFileName.FOURTH.toString());
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(10, counter)) + " %");
                                                            docLoad.saveFile(infoList,saveSampleDir);
                                                            loadStatus.setText("Загрузка: 100 %");
                                                            try {
                                                                docLoad.getClose();
                                                                xlxsOpen.getClose();
                                                            } catch (IOException e) {
                                                                e.printStackTrace();
                                                            }
                                                            if(mediumRangeOption){
                                                                docLoad.saveObrFile();
                                                            }
                                                        } catch (IOException | ClassNotFoundException e) {
                                                            e.printStackTrace();
                                                        }
                                                        counter_files++;
                                                    }
                                                }
                                                loadStatusFileNumber.setText("");
                                                loadStatus_end.setText("Успешно обработано " + counter_files + " файла(ов)!");
                                                checkStart = false;
                                                if(mediumRangeOption || genusOption || descriptionOption){
                                                    exceptionButton.setVisible(true);
                                                }
                                            }
                                        }.start();
                                    } else if(selectedSample.equals("Краткая версия микробиома")){
                                        new Thread(){
                                            @Override
                                            public void run(){
                                                counter_files = 0;
                                                for (int i = 0; i<samplePath.size();i++) {
                                                    if(samplePath.get(i).getPath().contains(".xlsx"))
                                                    {
                                                        loadStatusFileNumber.setText("Обработка " + i + " файла");
                                                        counter = 0;
                                                        infoList = new InfoList();
                                                        try {
                                                            xlxsOpen = new XLXSOpen(samplePath.get(i));
                                                            docLoad = new MainLoader("obr_2");
                                                            alg = new AlgOpen(infoList);
                                                        } catch (IOException e) {
                                                            e.printStackTrace();
                                                        } catch (InvalidFormatException e) {
                                                            e.printStackTrace();
                                                        }
                                                        try {
                                                            xlxsOpen.getPhylum(infoList);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(10, counter)) + " %");
                                                            xlxsOpen.getGenus(infoList);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(10, counter)) + " %");
                                                            xlxsOpen.getFileName(infoList);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(10, counter)) + " %");
                                                            xlxsOpen.getSpecies(infoList);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(10, counter)) + " %");
                                                            docLoad.setFileNameForThird(infoList);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(10, counter)) + " %");
                                                            docLoad.setFourTableFormatForSecond(infoList, 0);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(10, counter)) + " %");
                                                            docLoad.setAdditionForThird(infoList, 2);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(10, counter)) + " %");
                                                            //docLoad.saveSortedTable(infoList, 1, "Fifth");
                                                            //docLoad.saveSortedTable(infoList, 2, "Sixth");
                                                            docLoad.setTwoFormatWithSer(infoList, 1, SerFileName.FIFTH.toString());
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(10, counter)) + " %");
                                                            docLoad.setTwoFormatWithSer(infoList, 2, SerFileName.SIXTH.toString());
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(10, counter)) + " %");
                                                            docLoad.saveFile(infoList, saveSampleDir);
                                                            loadStatus.setText("Загрузка: 100 %");
                                                            try {
                                                                docLoad.getClose();
                                                                xlxsOpen.getClose();
                                                            } catch (IOException e) {
                                                                e.printStackTrace();
                                                            }
                                                            if(mediumRangeOption){
                                                                docLoad.saveObrFile();
                                                            }
                                                        } catch (IOException | ClassNotFoundException e) {
                                                            e.printStackTrace();
                                                        }
                                                        counter_files++;
                                                    }
                                                }
                                                loadStatusFileNumber.setText("");
                                                loadStatus_end.setText("Успешно обработано " + counter_files + " файла(ов)!");
                                                checkStart = false;
                                                if(mediumRangeOption || genusOption || descriptionOption){
                                                    exceptionButton.setVisible(true);
                                                }
                                            }
                                        }.start();
                                    } else if(selectedSample.equals("Образец взрослый стандартный")){
                                        new Thread(){
                                            @Override
                                            public void run(){
                                                counter_files = 0;
                                                for (int i = 0; i<samplePath.size();i++) {
                                                    if(samplePath.get(i).getPath().contains(".xlsx"))
                                                    {
                                                        loadStatusFileNumber.setText("Обработка " + i + " файла");
                                                        counter = 0;
                                                        infoList = new InfoList();
                                                        try {
                                                            xlxsOpen = new XLXSOpen(samplePath.get(i));
                                                            docLoad = new MainLoader("obr_3");
                                                            alg = new AlgOpen(infoList);
                                                        } catch (IOException e) {
                                                            e.printStackTrace();
                                                        } catch (InvalidFormatException e) {
                                                            e.printStackTrace();
                                                        }
                                                        try {
                                                            xlxsOpen.getPhylum(infoList);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(20, counter)) + " %");
                                                            xlxsOpen.getGenus(infoList);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(20, counter)) + " %");
                                                            xlxsOpen.getFileName(infoList);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(20, counter)) + " %");
                                                            xlxsOpen.getSpecies(infoList);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(20, counter)) + " %");
                                                            docLoad.setFileNameForFirst(infoList);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(20, counter)) + " %");
                                                            docLoad.setFourFormat(infoList, 2);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(20, counter)) + " %");
                                                            docLoad.setFiveFormat(infoList, 3);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(20, counter)) + " %");
                                                            docLoad.setFiveFormat(infoList, 4);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(20, counter)) + " %");
                                                            docLoad.setFourFormat(infoList, 5);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(20, counter)) + " %");
                                                            docLoad.setFourFormat(infoList, 6);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(20, counter)) + " %");
                                                            docLoad.setFourFormat(infoList, 7);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(20, counter)) + " %");
                                                            docLoad.setFourFormat(infoList, 8);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(20, counter)) + " %");
                                                            docLoad.setFourFormat(infoList, 9);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(20, counter)) + " %");
                                                            docLoad.setFourFormat(infoList, 10);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(20, counter)) + " %");
                                                            docLoad.setFourFormat(infoList, 11);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(20, counter)) + " %");
                                                            docLoad.setFourFormat(infoList, 12);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(20, counter)) + " %");
                                                            docLoad.setAddition(infoList);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(20, counter)) + " %");
                                                            //docLoad.saveSortedTable(infoList, 14, "Seventh");
                                                            //docLoad.saveSortedTable(infoList, 15, "Eighth");
                                                            docLoad.setTwoFormatWithSer(infoList, 14, SerFileName.SEVENTH.toString());
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(20, counter)) + " %");
                                                            docLoad.setTwoFormatWithSer(infoList, 15, SerFileName.EIGHTH.toString());
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(20, counter)) + " %");
                                                            docLoad.saveFile(infoList, saveSampleDir);
                                                            loadStatus.setText("Загрузка: 100 %");
                                                            try {
                                                                docLoad.getClose();
                                                                xlxsOpen.getClose();
                                                            } catch (IOException e) {
                                                                e.printStackTrace();
                                                            }
                                                            if(mediumRangeOption){
                                                                docLoad.saveObrFile();
                                                            }
                                                        } catch (IOException | ClassNotFoundException e) {
                                                            e.printStackTrace();
                                                        }
                                                        counter_files++;
                                                    }
                                                }
                                                loadStatusFileNumber.setText("");
                                                loadStatus_end.setText("Успешно обработано " + counter_files + " файла(ов)!");
                                                checkStart = false;
                                                if(mediumRangeOption || genusOption || descriptionOption){
                                                    exceptionButton.setVisible(true);
                                                }
                                            }
                                        }.start();
                                    } else if(selectedSample.equals("Микробиом кишечника")){
                                        new Thread(){
                                            @Override
                                            public void run(){
                                                counter_files = 0;
                                                for (int i = 0; i<samplePath.size();i++) {
                                                    if(samplePath.get(i).getPath().contains(".xlsx"))
                                                    {
                                                        loadStatusFileNumber.setText("Обработка " + i + " файла");
                                                        counter = 0;
                                                        infoList = new InfoList();
                                                        try {
                                                            xlxsOpen = new XLXSOpen(samplePath.get(i));
                                                            docLoad = new MainLoader("obr_4");
                                                            alg = new AlgOpen(infoList);
                                                        } catch (IOException e) {
                                                            e.printStackTrace();
                                                        } catch (InvalidFormatException e) {
                                                            e.printStackTrace();
                                                        }
                                                        try {
                                                            xlxsOpen.getPhylum(infoList);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(15, counter)) + " %");
                                                            xlxsOpen.getGenus(infoList);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(15, counter)) + " %");
                                                            xlxsOpen.getFileName(infoList);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(15, counter)) + " %");
                                                            xlxsOpen.getSpecies(infoList);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(15, counter)) + " %");
                                                            docLoad.setFileNameForFifth(infoList);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(15, counter)) + " %");
                                                            docLoad.setFourFormat(infoList, 2);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(15, counter)) + " %");
                                                            docLoad.setFourFormat(infoList, 3);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(15, counter)) + " %");
                                                            docLoad.setFiveFormat(infoList, 4);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(15, counter)) + " %");
                                                            docLoad.setFourFormat(infoList, 5);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(15, counter)) + " %");
                                                            docLoad.setFourFormat(infoList, 6);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(15, counter)) + " %");
                                                            docLoad.setFourFormat(infoList, 7);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(15, counter)) + " %");
                                                            docLoad.setAddition(infoList);
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(15, counter)) + " %");
                                                            //docLoad.saveSortedTable(infoList, 9, "Ninth");
                                                            //docLoad.saveSortedTable(infoList, 10, "Tenth");
                                                            docLoad.setTwoFormatWithSer(infoList, 9, SerFileName.NINTH.toString());
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(15, counter)) + " %");
                                                            docLoad.setTwoFormatWithSer(infoList, 10, SerFileName.TENTH.toString());
                                                            loadStatus.setText("Загрузка: " + (counter=getCounter(15, counter)) + " %");
                                                            docLoad.saveFile(infoList, saveSampleDir);
                                                            loadStatus.setText("Загрузка: 100 %");
                                                            try {
                                                                docLoad.getClose();
                                                                xlxsOpen.getClose();
                                                            } catch (IOException e) {
                                                                e.printStackTrace();
                                                            }
                                                            if(mediumRangeOption){
                                                                docLoad.saveObrFile();
                                                            }
                                                        } catch (IOException | ClassNotFoundException e) {
                                                            e.printStackTrace();
                                                        }
                                                        counter_files++;
                                                    }
                                                }
                                                loadStatusFileNumber.setText("");
                                                loadStatus_end.setText("Успешно обработано " + counter_files + " файла(ов)!");
                                                checkStart = false;
                                                if(mediumRangeOption || genusOption || descriptionOption){
                                                    exceptionButton.setVisible(true);
                                                }
                                            }
                                        }.start();
                                    }
                                } else
                                {
                                    errorMessageStr = "Выбранная папка загрузки является пустой...";
                                    ErrorController errorController = new ErrorController();
                                    try {
                                        errorController.start(new Stage());
                                    } catch (IOException e) {
                                        e.printStackTrace();
                                    }
                                }
                            } else {
                                errorMessageStr = "Вы не выбрали шаблон для создания отчета...";
                                ErrorController errorController = new ErrorController();
                                try {
                                    errorController.start(new Stage());
                                } catch (IOException e) {
                                    e.printStackTrace();
                                }
                            }
                        } else {
                            errorMessageStr = "Вы не указаали директорию загрузки или директорию выгрузки...";
                            ErrorController errorController = new ErrorController();
                            try {
                                errorController.start(new Stage());
                            } catch (IOException e) {
                                e.printStackTrace();
                            }
                        }
                    } else
                    {
                        errorMessageStr = "Происходит обработка файлов. Повторите попытку попытку позже...";
                        ErrorController errorController = new ErrorController();
                        try {
                            errorController.start(new Stage());
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    }
                }
        );
    }
}
