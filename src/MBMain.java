import java.io.File;
import java.io.FileNotFoundException;
import java.util.concurrent.TimeUnit;

import javafx.application.Application;
import javafx.geometry.HPos;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.TextField;
import javafx.scene.layout.Border;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.BorderStroke;
import javafx.scene.layout.BorderStrokeStyle;
import javafx.scene.layout.BorderWidths;
import javafx.scene.layout.ColumnConstraints;
import javafx.scene.layout.CornerRadii;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
import javafx.scene.paint.Color;
import javafx.scene.text.Text;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

public class MBMain extends Application {

	//Source Spreadsheet 
	private File sourceSS;
	private FileChooser sourceSSFC = new FileChooser();
	private Text sourceSSLbl = new Text("Spreadsheet: ");
	//Dest Folder
	private File destDir;
	private DirectoryChooser destDirDC = new DirectoryChooser();
	private Text destDirLbl = new Text("Destination folder: ");
	//Message Board 
	private Text msgText = new Text();
	//Enter Button 
	private Button go = new Button("Go");

	
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		launch();
	}

	@Override
	public void start(Stage primaryStage) throws Exception {
		
		// Application Title
		primaryStage.setTitle("Balances Formatter");
				
		//main layout
		BorderPane mainLayout = new BorderPane();
		
		// Source Panel -> Top 
//		VBox srcPanel =  makeSrcPanel(primaryStage); 
		GridPane  mainPanel =  makeMainPanel(primaryStage); 
		mainLayout.setCenter(mainPanel);
		
		VBox msgBoard = makeMSGPanel(primaryStage);
		mainLayout.setBottom(msgBoard);
		
		Scene mainScene = new Scene(mainLayout);
		primaryStage.setScene(mainScene);
		primaryStage.show();
		
	}
	
private GridPane makeMainPanel(Stage primaryStage){
		
		// Main Panel
		GridPane mainPanel = new GridPane();
		mainPanel.setHgap(10);
		mainPanel.setVgap(12);
		mainPanel.setAlignment(Pos.CENTER);
		mainPanel.setPadding(new Insets(15,15,15,15));
		
		ColumnConstraints column1 = new ColumnConstraints();
		column1.setHalignment(HPos.RIGHT);
		mainPanel.getColumnConstraints().add(column1);
		
		// Source Spreadsheet
		TextField srcSSTF = new TextField();
		srcSSTF.setPrefWidth(400);
		Button srcBTN = new Button("~");
		srcBTN.setPrefSize(20, 20);
		srcBTN.setOnAction(e -> {
			sourceSS = sourceSSFC.showOpenDialog(primaryStage);
			srcSSTF.setText(sourceSS.getAbsolutePath());
	    });
		
		mainPanel.add(sourceSSLbl,0,0);
		mainPanel.add(srcSSTF,1,0,2,1);
		mainPanel.add(srcBTN,3,0);
		
		// dest directory
		TextField destDirTF = new TextField();
		destDirTF.setPrefWidth(400);
		Button destBTN = new Button("~");
		destBTN.setPrefSize(20, 20);
		destBTN.setOnAction(e -> {
			destDir = destDirDC.showDialog(primaryStage);
			destDirTF.setText(destDir.getAbsolutePath());
	    });
		
		mainPanel.add(destDirLbl,0,1);
		mainPanel.add(destDirTF,1,1,2,1);
		mainPanel.add(destBTN,3,1);
		
		// Select Month
		TextField monthTF = new TextField();
		monthTF.setPrefWidth(200);
		Text monthLbl = new Text("Month: ");
		
		mainPanel.add(monthLbl,0,2);
		mainPanel.add(monthTF,1,2,1,1);	
		
		go.setPrefWidth(100);
		mainPanel.add(go,2,3);
		mainPanel.setHalignment(go, HPos.RIGHT);
		
		go.setOnAction(value -> {
//			File ab = new File("bin/Clients.xlsx");
//			String b = ab.getAbsolutePath();
//			"/Users/ricardochejfec/Programming/"
			try {
				MBDocMaker a = new MBDocMaker(sourceSS.getAbsolutePath(), destDir.getAbsolutePath(), monthTF.getText());
				String status = a.createDoc();
				msgText.setText(status);
			}
			catch (FileNotFoundException e){
				msgText.setText(e.getMessage());
			}
			catch (Exception e){
				msgText.setText(e.getMessage());
			}
			//wait x ammount 
			//then exit
			
			
			
		});
		
		return mainPanel;
	}

private VBox makeMSGPanel(Stage primaryStage){
	
	VBox msgPanel = new VBox();
	
	msgPanel.setPadding(new Insets(15,15,15,15));
	msgPanel.setSpacing(15);
	msgPanel.setAlignment(Pos.CENTER);
	msgPanel.setPrefSize(100, 100);
	msgText.setText("Ready to use.");
	
	msgPanel.getChildren().add(msgText);
	
	
	msgPanel.setBorder(new Border(new BorderStroke(Color.GRAY, Color.GRAY, Color.GRAY, Color.GRAY,
            BorderStrokeStyle.DOTTED, BorderStrokeStyle.DOTTED, BorderStrokeStyle.DOTTED, BorderStrokeStyle.DOTTED,
            CornerRadii.EMPTY, new BorderWidths(1), Insets.EMPTY)));
	
	return msgPanel;	
}

}
