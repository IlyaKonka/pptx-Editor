package gui;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.image.Image;
import javafx.stage.Stage;


public class Main extends Application {


    @Override
    public void start(Stage primaryStage) throws Exception{


        Parent root = FXMLLoader.load(getClass().getResource("/sampleJava10.fxml"));//if java 10
        //Parent root = FXMLLoader.load(getClass().getResource("/sampleJava8.fxml"));//if java 8


        primaryStage.getIcons().add(new Image(getClass().getResource("/logo.png").toString()));

        Scene scene = new Scene(root,605,650); //if java 10
        //Scene scene = new Scene(root,593,700); //if java 8


        primaryStage.setResizable(false);
        primaryStage.setScene(scene);
        primaryStage.setTitle("pptxEditor.fx");
        primaryStage.show();

    }

    public static void main(String[] args) {
        launch(args);
    }

}
