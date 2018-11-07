package gui;

import javafx.application.Platform;
import javafx.beans.value.ChangeListener;
import javafx.beans.value.ObservableValue;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.Node;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.input.KeyCode;
import javafx.scene.input.KeyEvent;
import javafx.scene.input.MouseEvent;
import javafx.scene.layout.Priority;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import parser.PptxParser;

import javax.swing.*;
import java.awt.*;
import java.io.*;
import java.net.*;
import java.nio.file.FileSystems;
import java.nio.file.Path;
import java.util.ResourceBundle;


public class Controller implements Initializable {

    private boolean isFinished = false;
    private boolean isFinishedBackup = false;
    private boolean isReadyToStart = true;
    private boolean isReadyToStartBackup = true;
    private boolean isDirectoryChooserUsed = false;
    private boolean isDirectoryChooserUsedBackup = false;
    private PrintStream ps;
    private PrintStream psBackup;

    @FXML
    TextField bTextfield1;
    @FXML
    TextField bTextfield2;
    @FXML
    TextField bTextfield3;
    @FXML
    TextField bTextfield4;
    @FXML
    Button startButtonBackup;
    @FXML
    Button cleanButtonBackup;
    @FXML
    ProgressBar progressBackup;
    @FXML
    TextArea consoleBackup;
    @FXML
    ImageView imageBackup;
    @FXML
    ImageView pathBackupB;
    @FXML
    ImageView pathListB;


    @FXML
    TextField textfield1;
    @FXML
    TextField textfield2;
    @FXML
    TextField textfield3;
    @FXML
    TextField textfield4;
    @FXML
    TextField textfield5;
    @FXML
    TextField textfield6;
    @FXML
    TextField textfield7;
    @FXML
    TextField textfield8;
    @FXML
    TextField textfield9;
    @FXML
    CheckBox checkbox1;
    @FXML
    CheckBox checkbox2;
    @FXML
    CheckBox checkbox3;
    @FXML
    CheckBox checkbox4;
    @FXML
    Button startButton;
    @FXML
    Button cleanButton;
    @FXML
    ComboBox combobox1;
    @FXML
    ComboBox combobox2;
    @FXML
    TabPane tabPane;
    @FXML
    Hyperlink git;
    @FXML
    ImageView imageEye;
    @FXML
    ImageView pathPresMain;
    @FXML
    ImageView pathBackupMain;
    @FXML
    ImageView pathListMain;
    @FXML
    ProgressBar progress;
    @FXML
    TextArea console;


    public void initialize(URL location, ResourceBundle resources) {

        progress.setProgress(0.0);
        combobox1.getItems().addAll("Normal", "Bold", "Italic");
        combobox2.getItems().addAll("Black", "White", "Red", "Green", "Blue");
        Platform.runLater(() -> git.requestFocus());
        textfield4.setDisable(true);
        combobox1.setDisable(true);
        combobox2.setDisable(true);
        checkbox3.setDisable(true);
        checkbox4.setDisable(true);


        imageEye.setDisable(true);
        imageEye.setVisible(false);
        pathBackupMain.setDisable(true);
        pathBackupMain.setVisible(false);
        pathListMain.setDisable(true);
        pathListMain.setVisible(false);
        textfield5.setDisable(true);
        textfield6.setDisable(true);
        textfield7.setDisable(true);
        textfield8.setDisable(true);
        textfield9.setDisable(true);
        textfield9.setVisible(false);

        bTextfield4.setVisible(false);
        progressBackup.setProgress(0.0);

        git.setOnAction(e -> {
            if (Desktop.isDesktopSupported()) {
                try {
                    Desktop.getDesktop().browse(new URI("https://github.com/IlyaKonka/pptxEditor.git"));
                } catch (IOException e1) {
                    e1.printStackTrace();
                } catch (URISyntaxException e1) {
                    e1.printStackTrace();
                }
            }
        });


        console.textProperty().addListener(new ChangeListener<String>() {
            @Override
            public void changed(ObservableValue<? extends String> observable, String oldValue, String newValue) {

                if (isFinished = true) {
                    progress.setProgress(0);
                    isFinished = false;
                }
            }
        });


        consoleBackup.textProperty().addListener(new ChangeListener<String>() {
            @Override
            public void changed(ObservableValue<? extends String> observable, String oldValue, String newValue) {
                if (isFinishedBackup = true) {
                    progressBackup.setProgress(0);
                    isFinishedBackup = false;
                }
            }
        });

        checkbox1.selectedProperty().addListener(new ChangeListener<Boolean>() {
            @Override
            public void changed(ObservableValue<? extends Boolean> observable, Boolean oldValue, Boolean newValue) {
                textfield4.setDisable(!checkbox1.isSelected());
                combobox1.setDisable(!checkbox1.isSelected());
                combobox2.setDisable(!checkbox1.isSelected());
                checkbox3.setDisable(!checkbox1.isSelected());
                checkbox4.setDisable(!checkbox1.isSelected());
            }
        });
        checkbox2.selectedProperty().addListener(new ChangeListener<Boolean>() {
            @Override
            public void changed(ObservableValue<? extends Boolean> observable, Boolean oldValue, Boolean newValue) {
                textfield5.setDisable(!checkbox2.isSelected());
                textfield6.setDisable(!checkbox2.isSelected());
                textfield7.setDisable(!checkbox2.isSelected());
                textfield8.setDisable(!checkbox2.isSelected());
                imageEye.setDisable(!checkbox2.isSelected());
                imageEye.setVisible(checkbox2.isSelected());
                pathBackupMain.setDisable(!checkbox2.isSelected());
                pathBackupMain.setVisible(checkbox2.isSelected());
                pathListMain.setDisable(!checkbox2.isSelected());
                pathListMain.setVisible(checkbox2.isSelected());
            }
        });

        imageEye.setPickOnBounds(true); // allows click on transparent areas
        imageEye.setOnMousePressed((MouseEvent e) -> {
            textfield6.setVisible(false);
            textfield6.setDisable(true);
            textfield9.setText(textfield6.getText());
            textfield9.setVisible(true);
            textfield9.setDisable(false);
        });

        imageEye.setOnMouseReleased((MouseEvent e) -> {
            textfield6.setVisible(true);
            textfield6.setDisable(false);
            textfield6.setText(textfield9.getText());
            textfield9.setVisible(false);
            textfield9.setDisable(true);
        });


        imageBackup.setPickOnBounds(true); // allows click on transparent areas
        imageBackup.setOnMousePressed((MouseEvent e) -> {
            bTextfield2.setVisible(false);
            bTextfield2.setDisable(true);
            bTextfield4.setText(bTextfield2.getText());
            bTextfield4.setVisible(true);
            bTextfield4.setDisable(false);
        });

        imageBackup.setOnMouseReleased((MouseEvent e) -> {
            bTextfield2.setVisible(true);
            bTextfield2.setDisable(false);
            bTextfield2.setText(bTextfield4.getText());
            bTextfield4.setVisible(false);
            bTextfield4.setDisable(true);
        });


        pathPresMain.setPickOnBounds(true); // allows click on transparent areas
        pathPresMain.setOnMouseClicked((MouseEvent e) -> {
            DirectoryChooser directoryChooser = new DirectoryChooser();
            Stage stage = (Stage) tabPane.getScene().getWindow();
            File file = directoryChooser.showDialog(stage);
            if (file != null)
                textfield1.setText(file.getAbsolutePath());
        });

        pathBackupMain.setPickOnBounds(true);
        pathBackupMain.setOnMouseClicked((MouseEvent e) -> {
            DirectoryChooser directoryChooser = new DirectoryChooser();
            Stage stage = (Stage) tabPane.getScene().getWindow();
            File file = directoryChooser.showDialog(stage);
            if (file != null) {
                isDirectoryChooserUsed = true;
                textfield5.setText(file.getAbsolutePath() + "\\backup.zip");
                textfield7.setText(file.getAbsolutePath() + "\\presentationList.txt");
            }
        });

        pathListMain.setPickOnBounds(true);
        pathListMain.setOnMouseClicked((MouseEvent e) -> {
            DirectoryChooser directoryChooser = new DirectoryChooser();
            Stage stage = (Stage) tabPane.getScene().getWindow();
            File file = directoryChooser.showDialog(stage);
            if (file != null) {
                isDirectoryChooserUsed = true;
                textfield7.setText(file.getAbsolutePath() + "\\presentationList.txt");
                textfield5.setText(file.getAbsolutePath() + "\\backup.zip");
            }
        });


        pathBackupB.setPickOnBounds(true);
        pathBackupB.setOnMouseClicked((MouseEvent e) -> {
            DirectoryChooser directoryChooser = new DirectoryChooser();
            Stage stage = (Stage) tabPane.getScene().getWindow();
            File file = directoryChooser.showDialog(stage);
            if (file != null) {
                isDirectoryChooserUsedBackup = true;
                bTextfield1.setText(file.getAbsolutePath() + "\\backup.zip");
                bTextfield3.setText(file.getAbsolutePath() + "\\presentationList.txt");
            }
        });

        pathListB.setPickOnBounds(true);
        pathListB.setOnMouseClicked((MouseEvent e) -> {
            DirectoryChooser directoryChooser = new DirectoryChooser();
            Stage stage = (Stage) tabPane.getScene().getWindow();
            File file = directoryChooser.showDialog(stage);
            if (file != null) {
                isDirectoryChooserUsedBackup = true;
                bTextfield3.setText(file.getAbsolutePath() + "\\presentationList.txt");
                bTextfield1.setText(file.getAbsolutePath() + "\\backup.zip");
            }
        });


        //isDirectoryChooserUsedBackup


        new EventHandler<KeyEvent>() {
            public void handle(KeyEvent event) {

                KeyCode code = event.getCode();
                if (code == KeyCode.TAB && !event.isShiftDown() && !event.isControlDown()) {
                    event.consume();
                    Node node = (Node) event.getSource();
                    KeyEvent newEvent
                            = new KeyEvent(event.getSource(),
                            event.getTarget(), event.getEventType(),
                            event.getCharacter(), event.getText(),
                            event.getCode(), event.isShiftDown(),
                            true, event.isAltDown(),
                            event.isMetaDown());

                    node.fireEvent(newEvent);
                }
            }
        };


        textfield5.setOnKeyPressed(new EventHandler<KeyEvent>() {
            @Override
            public void handle(KeyEvent keyEvent) {
                KeyCode code = keyEvent.getCode();

                String text = textfield5.getText();
                if ((text.endsWith(".zip") || text.endsWith(".zi"))  && !isDirectoryChooserUsed) {
                    if (text.endsWith(".zip")) {
                        text = text.substring(0, text.lastIndexOf(".zip"));
                    } else {
                        text = text.substring(0, text.lastIndexOf(".zi"));
                    }
                    char[] textChars = text.toCharArray();
                    for (int i = text.length()-1; i >= 0; i--) {
                        if(!Character.isLetter(textChars[i]))
                            text=text.substring(0,i+1);
                    }
                    textfield7.setText(text + "presentationsList.txt");
                }

            }
        });

        bTextfield1.setOnKeyPressed(new EventHandler<KeyEvent>() {
            @Override
            public void handle(KeyEvent keyEvent) {
                KeyCode code = keyEvent.getCode();

                String text = bTextfield1.getText();
                if ((text.endsWith(".zip") || text.endsWith(".zi")) && !isDirectoryChooserUsedBackup) {
                    if (text.endsWith(".zip")) {
                        text = text.substring(0, text.lastIndexOf(".zip"));
                    } else {
                        text = text.substring(0, text.lastIndexOf(".zi"));
                    }
                    char[] textChars = text.toCharArray();
                    for (int i = text.length()-1; i >= 0; i--) {
                        if(!Character.isLetter(textChars[i]))
                            text=text.substring(0,i+1);
                    }
                    bTextfield3.setText(text + "presentationsList.txt");
                }
            }
        });
    }


    private void setDisable(boolean isDisable) {
        textfield1.setDisable(isDisable);
        textfield2.setDisable(isDisable);
        textfield3.setDisable(isDisable);
        textfield4.setDisable(isDisable);
        textfield5.setDisable(isDisable);
        textfield6.setDisable(isDisable);
        textfield7.setDisable(isDisable);
        textfield8.setDisable(isDisable);
        combobox1.setDisable(isDisable);
        combobox2.setDisable(isDisable);
        checkbox2.setDisable(isDisable);
        checkbox3.setDisable(isDisable);
        checkbox4.setDisable(isDisable);
        checkbox1.setDisable(isDisable);
        imageEye.setDisable(true);
        imageEye.setVisible(false);
        pathPresMain.setDisable(true);
        pathPresMain.setVisible(false);
        pathBackupMain.setDisable(true);
        pathBackupMain.setVisible(false);
        pathListMain.setDisable(true);
        pathListMain.setVisible(false);
        textfield9.setDisable(true);
        textfield9.setVisible(false);
        startButton.setDisable(isDisable);
    }


    public void pressStart(ActionEvent event) throws InterruptedException {

        progress.setProgress(-1);
        Thread parseThread = new Thread(new Runnable() {
            @Override
            public void run() {
                ps = new PrintStream(new Console(console));
                System.setOut(ps);
                System.setErr(ps);
                setDisable(true);
                String presentationPath = textfield1.getText();
                String oldWord = textfield2.getText();
                String newWord = textfield3.getText();
                String pictureText = null;
                boolean pictureMode = false;
                String style = "normal";
                String color = "black";
                boolean manyTimesWrite = false;
                boolean toRotate = false;
                boolean toSave = false;
                String zipFolder = textfield5.getText();
                String zipPassword = null;
                zipPassword = textfield6.getText();
                String presentationListFolder = textfield7.getText();
                String receiverMailAdr = textfield8.getText();


                if (presentationPath == null || oldWord == null || newWord == null
                        || presentationPath.trim().isEmpty() || oldWord.trim().isEmpty() || newWord.trim().isEmpty()) {
                    System.out.println("Please, write:\nPath to presentation(s)\nWord(s) for replace\nNew word(s)\n");
                    isReadyToStart = false;
                }
                if (checkbox1.isSelected()) //picture Mode
                {
                    pictureMode = true;
                    pictureText = textfield4.getText();
                    if (pictureText == null || pictureText.trim().isEmpty())
                        pictureText = null;

                    if (combobox1.getSelectionModel().getSelectedIndex() == -1
                            || combobox1.getSelectionModel().getSelectedIndex() == 0)
                        style = "normal";
                    else if (combobox1.getSelectionModel().getSelectedIndex() == 1)
                        style = "bold";
                    else if (combobox1.getSelectionModel().getSelectedIndex() == 2)
                        style = "italic";

                    if (combobox2.getSelectionModel().getSelectedIndex() == -1
                            || combobox2.getSelectionModel().getSelectedIndex() == 0)
                        color = "black";
                    else if (combobox2.getSelectionModel().getSelectedIndex() == 1)
                        color = "white";
                    else if (combobox2.getSelectionModel().getSelectedIndex() == 2)
                        color = "red";
                    else if (combobox2.getSelectionModel().getSelectedIndex() == 2)
                        color = "green";
                    else if (combobox2.getSelectionModel().getSelectedIndex() == 2)
                        color = "blue";

                    if (checkbox3.isSelected())
                        manyTimesWrite = true;

                    if (checkbox4.isSelected())
                        toRotate = true;
                }

                if (checkbox2.isSelected()) //backup Mode
                {
                    toSave = true;
                    if (zipFolder == null || zipPassword == null || presentationListFolder == null
                            || zipFolder.trim().isEmpty() || zipPassword.trim().isEmpty() || presentationListFolder.trim().isEmpty()) {
                        System.out.println("Please, write:\nPath to backup zip\nZip password\nPath to file with names of presentation\nOR\nDeselect \"backup\" checkbox");
                        isReadyToStart = false;
                    }

                    if (receiverMailAdr == null || receiverMailAdr.trim().isEmpty())
                        System.out.println("WARNING: No mail with information will be sent.");
                }

                if (isReadyToStart) {
                    PptxParser.parse(presentationPath, oldWord, newWord, pictureMode, pictureText, style, color, manyTimesWrite,
                            toRotate, toSave, zipFolder, zipPassword, presentationListFolder, receiverMailAdr);
                    isFinished = true;
                    isDirectoryChooserUsed = false;
                } else
                    isReadyToStart = true;

                ps.close();
            }
        });


        parseThread.start();


    }

    @FXML
    public void pressClean(ActionEvent event) {
        setDisable(false);
        textfield4.setDisable(!checkbox1.isSelected());
        combobox1.setDisable(!checkbox1.isSelected());
        combobox2.setDisable(!checkbox1.isSelected());
        checkbox3.setDisable(!checkbox1.isSelected());
        checkbox4.setDisable(!checkbox1.isSelected());
        imageEye.setDisable(true);
        imageEye.setVisible(false);
        pathListMain.setDisable(true);
        pathListMain.setVisible(false);
        pathBackupMain.setDisable(true);
        pathBackupMain.setVisible(false);
        textfield9.setDisable(true);
        textfield9.setVisible(false);
        textfield5.setDisable(!checkbox2.isSelected());
        textfield6.setDisable(!checkbox2.isSelected());
        textfield7.setDisable(!checkbox2.isSelected());
        textfield8.setDisable(!checkbox2.isSelected());
        progress.setProgress(0);
        console.clear();
        pathPresMain.setVisible(true);
        pathPresMain.setDisable(false);
        //combobox1.cancelEdit();
        //combobox2.cancelEdit();
        checkbox1.setSelected(false);
        checkbox2.setSelected(false);
        checkbox3.setSelected(false);
        checkbox4.setSelected(false);
    }

    private void setDisableBackup(boolean isDisable) {
        bTextfield1.setDisable(isDisable);
        bTextfield2.setDisable(isDisable);
        bTextfield3.setDisable(isDisable);
        imageBackup.setDisable(true);
        imageBackup.setVisible(false);
        pathBackupB.setDisable(true);
        pathBackupB.setVisible(false);
        pathListB.setDisable(true);
        pathListB.setVisible(false);
        bTextfield4.setDisable(true);
        bTextfield4.setVisible(false);
        startButtonBackup.setDisable(isDisable);
    }


    @FXML
    public void pressStartBackup(ActionEvent event) {
        progressBackup.setProgress(-1);
        Thread parseThread = new Thread(new Runnable() {
            @Override
            public void run() {
                psBackup = new PrintStream(new Console(consoleBackup));
                System.setOut(psBackup);
                System.setErr(psBackup);
                setDisableBackup(true);
                String zipFolder = bTextfield1.getText();
                String zipPassword = null;
                zipPassword = bTextfield2.getText();
                String presentationListFolder = bTextfield3.getText();


                if (zipFolder == null || zipPassword == null || presentationListFolder == null
                        || zipFolder.trim().isEmpty() || zipPassword.trim().isEmpty() || presentationListFolder.trim().isEmpty()) {
                    System.out.println("Please, write:\nPath to backup zip\nZip password\nPath to file with names of presentation\n");
                    isReadyToStartBackup = false;
                }


                if (isReadyToStartBackup) {
                    PptxParser.backup(zipFolder, zipPassword, presentationListFolder);
                    isFinishedBackup = true;
                    isDirectoryChooserUsedBackup = false;
                } else
                    isReadyToStartBackup = true;

                psBackup.close();
            }
        });

        parseThread.start();

    }

    @FXML
    public void pressCleanBackup(ActionEvent event) {
        setDisableBackup(false);
        imageBackup.setDisable(false);
        imageBackup.setVisible(true);
        pathListB.setDisable(false);
        pathListB.setVisible(true);
        pathBackupB.setDisable(false);
        pathBackupB.setVisible(true);
        bTextfield4.setDisable(true);
        bTextfield4.setVisible(false);
        progress.setProgress(0);
        consoleBackup.clear();
    }

    public class Console extends OutputStream {
        private TextArea cons;

        public Console(TextArea console) {
            this.cons = console;
        }

        public void appendText(String valueOf) {
            Platform.runLater(() -> cons.appendText(valueOf));
        }

        public void write(int b) throws IOException {
            appendText(String.valueOf((char) b));
        }
    }

}
