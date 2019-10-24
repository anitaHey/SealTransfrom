import javafx.scene.text.Font;

import javafx.scene.paint.Color;

import java.awt.Image;
import java.awt.Toolkit;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.Transferable;
import java.awt.datatransfer.UnsupportedFlavorException;
import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Optional;
import javax.imageio.ImageIO;

import javafx.application.Platform;
import javafx.application.Application;
import javafx.beans.property.SimpleStringProperty;
import javafx.beans.value.ChangeListener;
import javafx.beans.value.ObservableValue;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.HBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import javafx.util.Callback;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;
import javafx.scene.control.*;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.layout.Pane;
import javafx.scene.shape.Line;
import javafx.scene.shape.Circle;
import javafx.scene.image.WritableImage;
import javafx.scene.SnapshotParameters;
import javafx.embed.swing.SwingFXUtils;

public class Seal extends Application {
	public static void main(String[] args) {
		launch(args);
	}

	public void start(Stage primaryStage) throws IOException {

		Scene scene = new Scene(new StackPane());
		FXMLLoader loader = new FXMLLoader(getClass().getResource("Main.fxml"));
		scene.setRoot(loader.load());
		primaryStage.setTitle("印章生產器");
		primaryStage.setScene(scene);
		primaryStage.show();

	}

	@FXML
	Pane seal, seal_img;
	@FXML
	Button name_before, seal_before, save_copy, save_pic, enter, enter2;
	@FXML
	TextField name1, name2, word;
	@FXML
	ChoiceBox<String> font, choose, date_format;
	@FXML
	ColorPicker color;
	@FXML
	RadioButton choose_date, choose_word;
	@FXML
	DatePicker date;
	@FXML
	Circle circle;
	@FXML
	Line line_up, line_down;
	@FXML
	Label label_other = new Label(""), label_name1 = new Label(""), label_name2 = new Label(""), date1, date2;
	@FXML
	HBox hbox1 = new HBox(), hbox2 = new HBox(), hbox3 = new HBox();
	ToggleGroup group;
	ObservableList<data> all = FXCollections.observableArrayList();
	ObservableList<data_name> all_name = FXCollections.observableArrayList();
	double width1 = label_name1.getLayoutBounds().getWidth();
	double width2 = label_name2.getLayoutBounds().getWidth();
	double width3 = label_other.getLayoutBounds().getWidth();
	double font_size;

	public void initialize() {
		circle.setFill(Color.TRANSPARENT);
		ArrayList<String> families = new ArrayList<>();
		File file = new File("font");
		File[] list = file.listFiles();
		for (int a = 0; a < list.length; a++) {
			String tem = list[a].getName().substring(0, list[a].getName().length() - 4);
			families.add(tem);
		}
		font.getItems().addAll(families);
		font.setMaxWidth(Double.MAX_VALUE);
		date.setValue(NOW_LOCAL_DATE());
		date_format.getItems().addAll("yyyy.MM.dd", "yyyy/MM/dd", "\"yy.MM.dd", "\"yy/MM/dd", "yy.MM.dd", "yy/MM/dd");
		date_format.setValue(date_format.getItems().get(0));
		font.setValue("System");

		group = new ToggleGroup();
		choose_date.setSelected(true);
		choose_date.setToggleGroup(group);
		choose_word.setToggleGroup(group);
		
		TableView<data_name> table1 = new TableView<>();
		TableColumn<data_name, String> table_up1 = new TableColumn<>("上欄");
		table_up1.setCellValueFactory(new PropertyValueFactory<data_name, String>("up"));
		table_up1.prefWidthProperty().bind(table1.widthProperty().multiply(0.46));
		TableColumn<data_name, String> table_down1 = new TableColumn<>("下欄");
		table_down1.setCellValueFactory(new PropertyValueFactory<data_name, String>("down"));
		table_down1.prefWidthProperty().bind(table1.widthProperty().multiply(0.46));

		table1.getColumns().addAll(table_up1, table_down1);

		File file_open1 = new File("name.xls");
		Workbook workbook1 = null;
		try {
			workbook1 = Workbook.getWorkbook(file_open1);
		} catch (BiffException | IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		all_name.clear();
		Sheet readSheet1 = workbook1.getSheet(0);
		for (int a = 0; a < readSheet1.getRows(); a++) {
			String up1 = readSheet1.getCell(0, a).getContents();
			String down1 = readSheet1.getCell(1, a).getContents();

			all_name.add(new data_name(up1, down1));
		}
		workbook1.close();
		table1.setItems(all_name);
		
		
		
		TableView<data> table2 = new TableView<>();
		TableColumn<data, String> table_up2 = new TableColumn<>("上欄");
		table_up2.setCellValueFactory(new PropertyValueFactory<data, String>("up"));
		table_up2.prefWidthProperty().bind(table2.widthProperty().multiply(0.15));
		TableColumn<data, String> table_down2 = new TableColumn<>("下欄");
		table_down2.setCellValueFactory(new PropertyValueFactory<data, String>("down"));
		TableColumn<data, String> table_color2 = new TableColumn<>("顏色");
		table_down2.prefWidthProperty().bind(table2.widthProperty().multiply(0.13));
		table_color2.setCellValueFactory(new PropertyValueFactory<data, String>("color"));
		TableColumn<data, String> table_other2 = new TableColumn<>("中欄內容");
		table_color2.prefWidthProperty().bind(table2.widthProperty().multiply(0.16));
		table_other2.setCellValueFactory(new PropertyValueFactory<data, String>("other"));
		TableColumn<data, String> table_format2 = new TableColumn<>("日期格式");
		table_other2.prefWidthProperty().bind(table2.widthProperty().multiply(0.15));
		table_format2.setCellValueFactory(new PropertyValueFactory<data, String>("format"));
		TableColumn<data, String> table_word2 = new TableColumn<>("文字內容");
		table_format2.prefWidthProperty().bind(table2.widthProperty().multiply(0.17));
		table_word2.setCellValueFactory(new PropertyValueFactory<data, String>("word"));
		table_word2.prefWidthProperty().bind(table2.widthProperty().multiply(0.2));
		table_color2.setCellFactory(new Callback<TableColumn<data, String>, TableCell<data, String>>() {
			@Override
			public TableCell<data, String> call(TableColumn<data, String> p) {
				return new TableCell<data, String>() {
					@Override
					public void updateItem(final String item, final boolean empty) {
						super.updateItem(item, empty);// *don't forget!
						if (item != null) {
							setText(" ");
							setStyle("-fx-background-color:" + toRgbString(Color.web(item)));
						} else {
							setText(null);
						}
					}
				};
			}
		});
		table2.getColumns().addAll(table_up2, table_down2, table_color2, table_other2, table_format2, table_word2);

		File file_open2 = new File("data.xls");
		Workbook workbook2 = null;
		try {
			workbook2 = Workbook.getWorkbook(file_open2);
		} catch (BiffException | IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		all.clear();
		Sheet readSheet2 = workbook2.getSheet(0);
		for (int a = 0; a < readSheet2.getRows(); a++) {
			String up2 = readSheet2.getCell(0, a).getContents();
			String down2 = readSheet2.getCell(1, a).getContents();
			String color2 = readSheet2.getCell(2, a).getContents();
			String other2 = readSheet2.getCell(3, a).getContents();
			String format2 = readSheet2.getCell(4, a).getContents();
			String word2 = readSheet2.getCell(5, a).getContents();

			all.add(new data(up2, down2, color2, other2, format2, word2));
		}
		workbook2.close();
		table2.setItems(all);
		
		
		
		
		
		choose_word.setOnAction(new EventHandler<ActionEvent>() {
			@Override
			public void handle(ActionEvent arg0) {
				// TODO Auto-generated method stub
				word.setVisible(true);
				date.setVisible(false);
				date_format.setValue("");
				date_format.setVisible(false);
				date2.setText("文字內容");
				date1.setVisible(false);
				label_other.setText("");
				enter2.setVisible(true);
			}
		});
		choose_date.setOnAction(new EventHandler<ActionEvent>() {
			@Override
			public void handle(ActionEvent arg0) {
				// TODO Auto-generated method stub
				date_format.setValue(date_format.getItems().get(0));
				word.setVisible(false);
				date.setVisible(true);
				date_format.setVisible(true);
				date2.setText("日期選擇");
				date1.setVisible(true);
				LocalDate day_change = date.getValue();
				DateTimeFormatter formatter_change = DateTimeFormatter
						.ofPattern(date_format.getSelectionModel().getSelectedItem());
				label_other.setText(day_change.format(formatter_change));
				enter2.setVisible(false);
			}
		});

		color.setValue(Color.BLACK);
		color.setOnAction(new EventHandler<ActionEvent>() {
			@Override
			public void handle(ActionEvent arg0) {
				// TODO Auto-generated method stub
				circle.setStroke(color.getValue());
				line_up.setStroke(color.getValue());
				line_down.setStroke(color.getValue());
				label_name1.setStyle(label_name1.getStyle() + "-fx-text-fill:" + toRgbString(color.getValue()) + ";");
				label_name2.setStyle(label_name2.getStyle() + "-fx-text-fill:" + toRgbString(color.getValue()) + ";");
				label_other.setStyle(label_other.getStyle() + "-fx-text-fill:" + toRgbString(color.getValue()) + ";");
			}
		});

		enter.setOnAction(new EventHandler<ActionEvent>() {
			@Override
			public void handle(ActionEvent event) {
				try {
					createNameExcel(new data_name(name1.getText(), name2.getText()));
				} catch (WriteException | IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				label_name1.setText(name1.getText());
				width1 = label_name1.getWidth();
				if (width1 > 99) {
					font_size = label_name1.getFont().getSize();
					double scalex = 1 - Double.valueOf(width1 - 95) / 200;
					label_name1.setStyle(
							label_name1.getStyle() + String.format("-fx-font-size: %dpx;", (int) (scalex * font_size)));
					width1 = label_name1.getWidth();
				} else if (label_name1.getText().length() < 5) {
					label_name1.setStyle(label_name1.getStyle() + "-fx-font-size: 22px;");
				}

				label_name2.setText(name2.getText());
				width2 = label_name2.getWidth();
				if (width2 > 99) {
					font_size = label_name2.getFont().getSize();
					double scalex = 1 - Double.valueOf(width2 - 90) / 200;
					label_name2.setStyle(
							label_name2.getStyle() + String.format("-fx-font-size: %dpx;", (int) (scalex * font_size)));
					width2 = label_name2.getWidth();
				} else if (label_name2.getText().length() < 5) {
					label_name2.setStyle(label_name2.getStyle() + "-fx-font-size: 22px;");
				}
			}
		});

		enter2.setOnAction(new EventHandler<ActionEvent>() {
			@Override
			public void handle(ActionEvent event) {
				label_other.setText(word.getText());
				width2 = label_other.getWidth();
				if (width2 > 114) {
					font_size = label_other.getFont().getSize();
					double scalex = 1 - Double.valueOf(width2 - 110) / 200;
					label_other.setStyle(
							label_other.getStyle() + String.format("-fx-font-size: %dpx;", (int) (scalex * font_size)));
					width2 = label_other.getWidth();
				} else if (label_other.getText().length() < 5) {
					label_other.setStyle(label_other.getStyle() + "-fx-font-size: 22px;");
				}
			}
		});

		font.getSelectionModel().selectedIndexProperty().addListener(new ChangeListener<Number>() {
			@Override
			public void changed(ObservableValue<? extends Number> arg0, Number oldNum, Number newNum) {
				// TODO Auto-generated method stub
				Font f = Font.loadFont(getClass().getResourceAsStream(list[(int) newNum].getPath()), 20);
				label_other.setFont(f);
				label_name1.setFont(f);
				label_name2.setFont(f);
			}
		});

		LocalDate day1 = date.getValue();
		DateTimeFormatter formatter = DateTimeFormatter.ofPattern(date_format.getSelectionModel().getSelectedItem());
		label_other.setText(day1.format(formatter));

		date_format.setOnAction((event) -> {
			DateTimeFormatter formatter_change = DateTimeFormatter
					.ofPattern(date_format.getSelectionModel().getSelectedItem());
			label_other.setText(day1.format(formatter_change));
		});
		date.setOnAction((event) -> {
			LocalDate day_change = date.getValue();
			DateTimeFormatter formatter_change = DateTimeFormatter
					.ofPattern(date_format.getSelectionModel().getSelectedItem());
			label_other.setText(day_change.format(formatter_change));
		});

		save_pic.setOnAction((event) -> {
			saveAsPng();
			try {
				createExcel(new data(label_name1.getText(), label_name2.getText(), color.getValue().toString(),
						((RadioButton) group.getSelectedToggle()).getText(), date_format.getValue(), word.getText()));
			} catch (WriteException | IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			final Alert alert = new Alert(AlertType.CONFIRMATION, "儲存成功!", ButtonType.YES);
			alert.setTitle("提示");
			alert.setHeaderText("");
			final Optional<ButtonType> opt = alert.showAndWait();
		});

		save_copy.setOnAction((event) -> {
			try {
				createExcel(new data(label_name1.getText(), label_name2.getText(), color.getValue().toString(),
						((RadioButton) group.getSelectedToggle()).getText(), date_format.getValue(), word.getText()));
			} catch (WriteException | IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			SnapshotParameters sp = new SnapshotParameters();
			sp.setFill(Color.TRANSPARENT);
			WritableImage image = seal_img.snapshot(sp, null);

			File file_img = new File("chart.png");
			try {
				ImageIO.write(SwingFXUtils.fromFXImage(image, null), "png", file_img);
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			Image pic = null;
			try {
				pic = ImageIO.read(file_img);
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			write(pic);

			final Alert alert = new Alert(AlertType.CONFIRMATION, "已複製!", ButtonType.YES);
			alert.setTitle("提示");
			alert.setHeaderText("");
			final Optional<ButtonType> opt = alert.showAndWait();
		});

		seal_before.setOnAction((event) -> {
			Dialog<Integer> dialog = new Dialog<Integer>();
			
			File file_open = new File("data.xls");
			Workbook workbook = null;
			try {
				workbook = Workbook.getWorkbook(file_open);
			} catch (BiffException | IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			all.clear();
			Sheet readSheet = workbook.getSheet(0);
			for (int a = 0; a < readSheet.getRows(); a++) {
				String up = readSheet.getCell(0, a).getContents();
				String down = readSheet.getCell(1, a).getContents();
				String color = readSheet.getCell(2, a).getContents();
				String other = readSheet.getCell(3, a).getContents();
				String format = readSheet.getCell(4, a).getContents();
				String word = readSheet.getCell(5, a).getContents();

				all.add(new data(up, down, color, other, format, word));
			}
			workbook.close();
			table2.setItems(all);

			DialogPane dialogPane = dialog.getDialogPane();
			dialogPane.getButtonTypes().addAll(ButtonType.OK, ButtonType.CANCEL);
			dialogPane.setContent(table2);
			dialog.getDialogPane().setMinWidth(600);
			dialog.setTitle("曾使用過的印章");
			dialog.setResultConverter((ButtonType button) -> {
				if (button == ButtonType.OK) {
					return table2.getSelectionModel().getSelectedIndex();
				}
				
				return null;
			});
			Optional<Integer> opt = dialog.showAndWait();
			opt.ifPresent((Integer results) -> {
				Workbook workbooks = null;
				try {
					workbooks = Workbook.getWorkbook(file_open);
				} catch (BiffException | IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}

				Sheet readSheets = workbooks.getSheet(0);
				String up = readSheets.getCell(0, results).getContents();
				String down = readSheets.getCell(1, results).getContents();
				String colors = readSheets.getCell(2, results).getContents();
				String other = readSheets.getCell(3, results).getContents();
				String format = readSheets.getCell(4, results).getContents();
				String words = readSheets.getCell(5, results).getContents();
				workbooks.close();

				Platform.runLater(() -> {
					
					name1.setText(up);
					name2.setText(down);
					word.setText(words);
					label_name1.setText(up);
					label_name2.setText(down);
					label_other.setText("");
					color.setValue(Color.web(colors));
					circle.setStroke(color.getValue());
					line_up.setStroke(color.getValue());
					line_down.setStroke(color.getValue());
					label_name1
							.setStyle(label_name1.getStyle() + "-fx-text-fill:" + toRgbString(color.getValue()) + ";");
					label_name2
							.setStyle(label_name2.getStyle() + "-fx-text-fill:" + toRgbString(color.getValue()) + ";");
					label_other
							.setStyle(label_other.getStyle() + "-fx-text-fill:" + toRgbString(color.getValue()) + ";");
					date_format.setValue(format);
					if (other.equals("日期")) {
						date.setValue(NOW_LOCAL_DATE());
						LocalDate day = NOW_LOCAL_DATE();
						DateTimeFormatter formatter_change = DateTimeFormatter
								.ofPattern(date_format.getSelectionModel().getSelectedItem());
						label_other.setText(day.format(formatter_change));
						choose_date.setSelected(true);
						date_format.setValue(date_format.getItems().get(0));
						word.setVisible(false);
						date.setVisible(true);
						date_format.setVisible(true);
						date2.setText("日期選擇");
						date1.setVisible(true);
						enter2.setVisible(false);
					} else {
						choose_word.setSelected(true);
						word.setVisible(true);
						date.setVisible(false);
						date_format.setValue("");
						date_format.setVisible(false);
						date2.setText("文字內容");
						date1.setVisible(false);
						label_other.setText(words);
						enter2.setVisible(true);
					}
				});
			});

		});

		name_before.setOnAction((event) -> {
			Dialog<Integer> dialog = new Dialog<Integer>();
			File file_open = new File("name.xls");
			Workbook workbook = null;
			try {
				workbook = Workbook.getWorkbook(file_open);
			} catch (BiffException | IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			all_name.clear();
			Sheet readSheet = workbook.getSheet(0);
			for (int a = 0; a < readSheet.getRows(); a++) {
				String up = readSheet.getCell(0, a).getContents();
				String down = readSheet.getCell(1, a).getContents();

				all_name.add(new data_name(up, down));
			}
			workbook.close();
			table1.setItems(all_name);
			DialogPane dialogPane = dialog.getDialogPane();

			dialogPane.getButtonTypes().addAll(ButtonType.OK, ButtonType.CANCEL);
			dialogPane.setContent(table1);;
			dialog.getDialogPane().setMinWidth(250);
			dialog.setTitle("曾使用過的姓名");
			dialog.setResultConverter((ButtonType button) -> {
				if (button == ButtonType.OK) {
					return table1.getSelectionModel().getSelectedIndex();
				}
			
				return null;
			});
			Optional<Integer> opt = dialog.showAndWait();
			opt.ifPresent((Integer results) -> {
				Workbook workbooks = null;
				try {
					workbooks = Workbook.getWorkbook(file_open);
				} catch (BiffException | IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}

				Sheet readSheets = workbooks.getSheet(0);
				String up = readSheets.getCell(0, results).getContents();
				String down = readSheets.getCell(1, results).getContents();
				workbooks.close();

				Platform.runLater(() -> {
					name1.setText(up);
					name2.setText(down);
					label_name1.setText(up);
					label_name2.setText(down);
				});

			});

		});
	}

	public static Image read() {
		Transferable t = Toolkit.getDefaultToolkit().getSystemClipboard().getContents(null);

		try {
			if (t != null && t.isDataFlavorSupported(DataFlavor.imageFlavor)) {
				Image image = (Image) t.getTransferData(DataFlavor.imageFlavor);
				return image;
			}
		} catch (Exception e) {
		}

		return null;
	}

	public static void write(Image image) {
		if (image == null)
			throw new IllegalArgumentException("Image can't be null");

		ImageTransferable transferable = new ImageTransferable(image);
		Toolkit.getDefaultToolkit().getSystemClipboard().setContents(transferable, null);
	}

	static class ImageTransferable implements Transferable {
		private Image image;

		public ImageTransferable(Image image) {
			this.image = image;
		}

		public Object getTransferData(DataFlavor flavor) throws UnsupportedFlavorException {
			if (isDataFlavorSupported(flavor)) {
				return image;
			} else {
				throw new UnsupportedFlavorException(flavor);
			}
		}

		public boolean isDataFlavorSupported(DataFlavor flavor) {
			return flavor == DataFlavor.imageFlavor;
		}

		public DataFlavor[] getTransferDataFlavors() {
			return new DataFlavor[] { DataFlavor.imageFlavor };
		}
	}

	public void saveAsPng() {
		SnapshotParameters sp = new SnapshotParameters();
		sp.setFill(Color.TRANSPARENT);
		WritableImage image = seal_img.snapshot(sp, null);

		FileChooser fileChooser = new FileChooser();
		FileChooser.ExtensionFilter extFilter = new FileChooser.ExtensionFilter("PNG files (*.png)", "*.png");
		fileChooser.getExtensionFilters().add(extFilter);
		fileChooser.setTitle("Save Seal");
		File file = fileChooser.showSaveDialog(null);

		try {
			ImageIO.write(SwingFXUtils.fromFXImage(image, null), "png", file);
		} catch (IOException e) {
			// TODO: handle exception here
			final Alert alert = new Alert(AlertType.CONFIRMATION, "儲存失敗! ", ButtonType.YES);
			alert.setTitle("提示");
			alert.setHeaderText("");
			final Optional<ButtonType> opt = alert.showAndWait();
		}
	}

	public static final LocalDate NOW_LOCAL_DATE() {
		String date = new SimpleDateFormat("yyyy.MM.dd").format(Calendar.getInstance().getTime());
		DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy.MM.dd");
		LocalDate localDate = LocalDate.parse(date, formatter);
		return localDate;
	}

	private String toRgbString(Color c) {
		return "rgb(" + to255Int(c.getRed()) + "," + to255Int(c.getGreen()) + "," + to255Int(c.getBlue()) + ")";
	}

	private int to255Int(double d) {
		return (int) (d * 255);
	}

	public static class data_name {
		private SimpleStringProperty up;
		private SimpleStringProperty down;

		public data_name(String up, String down) {
			this.up = new SimpleStringProperty(up);
			this.down = new SimpleStringProperty(down);
		}

		public String getUp() {
			return up.get();
		}

		public String getDown() {
			return down.get();
		}
	}

	public static class data {
		private SimpleStringProperty up;
		private SimpleStringProperty down;
		private SimpleStringProperty color;
		private SimpleStringProperty other;
		private SimpleStringProperty format;
		private SimpleStringProperty word;

		public data(String up, String down, String color, String other, String format, String word) {
			this.up = new SimpleStringProperty(up);
			this.down = new SimpleStringProperty(down);
			this.color = new SimpleStringProperty(color);
			this.other = new SimpleStringProperty(other);
			this.format = new SimpleStringProperty(format);
			this.word = new SimpleStringProperty(word);
		}

		public String getUp() {
			return up.get();
		}

		public String getDown() {
			return down.get();
		}

		public String getColor() {
			return color.get();
		}

		public String getOther() {
			return other.get();
		}

		public String getFormat() {
			return format.get();
		}

		public String getWord() {
			return word.get();
		}
	}
	
	public void createNameExcel(data_name new_data) throws IOException, RowsExceededException, WriteException {
		File file = new File("name.xls");
		Workbook workbook = null;
		try {
			workbook = Workbook.getWorkbook(file);
		} catch (BiffException | IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		all_name.clear();
		Sheet readSheet = workbook.getSheet(0);
		for (int a = 0; a < readSheet.getRows(); a++) {
			String up = readSheet.getCell(0, a).getContents();
			String down = readSheet.getCell(1, a).getContents();
			all_name.add(new data_name(up, down));
		}
		int check = 0;
		for (int a = 0; a < all_name.size(); a++) {
			if(all_name.get(a).getUp().equals(new_data.getUp())&&all_name.get(a).getDown().equals(new_data.getDown())) {
				check = 1;
			}
		}
		if(check!=1)
			all_name.add(new_data);
		workbook.close();
		WritableWorkbook writeWorkBook = Workbook.createWorkbook(file);
		WritableSheet writeSheet = writeWorkBook.createSheet("name", 0);

		for (int a = 0; a < all_name.size(); a++) {
			writeSheet.addCell(new jxl.write.Label(0, a, all_name.get(a).getUp()));
			writeSheet.addCell(new jxl.write.Label(1, a, all_name.get(a).getDown()));
		}
		
		writeWorkBook.write();
		writeWorkBook.close();


	}
	public void createExcel(data new_data) throws IOException, RowsExceededException, WriteException {
		File file = new File("data.xls");
		Workbook workbook = null;
		try {
			workbook = Workbook.getWorkbook(file);
		} catch (BiffException | IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		all.clear();
		Sheet readSheet = workbook.getSheet(0);
		for (int a = 0; a < readSheet.getRows(); a++) {
			String up = readSheet.getCell(0, a).getContents();
			String down = readSheet.getCell(1, a).getContents();
			String color = readSheet.getCell(2, a).getContents();
			String other = readSheet.getCell(3, a).getContents();
			String format = readSheet.getCell(4, a).getContents();
			String word = readSheet.getCell(5, a).getContents();

			all.add(new data(up, down, color, other, format, word));
		}
		int check = 0;
		for (int a = 0; a < all.size(); a++) {
			if(all.get(a).getUp().equals(new_data.getUp())&&all.get(a).getDown().equals(new_data.getDown())&&all.get(a).getColor().equals(new_data.getColor())&&all.get(a).getOther().equals(new_data.getOther())&&all.get(a).getFormat().equals(new_data.getFormat())&&all.get(a).getWord().equals(new_data.getWord())) {
				check = 1;
			}
		}
		if(check!=1)
			all.add(new_data);
		workbook.close();
		WritableWorkbook writeWorkBook = Workbook.createWorkbook(file);
		WritableSheet writeSheet = writeWorkBook.createSheet("data", 0);

		for (int a = 0; a < all.size(); a++) {
			writeSheet.addCell(new jxl.write.Label(0, a, all.get(a).getUp()));
			writeSheet.addCell(new jxl.write.Label(1, a, all.get(a).getDown()));
			writeSheet.addCell(new jxl.write.Label(2, a, all.get(a).getColor()));
			writeSheet.addCell(new jxl.write.Label(3, a, all.get(a).getOther()));
			writeSheet.addCell(new jxl.write.Label(4, a, all.get(a).getFormat()));
			writeSheet.addCell(new jxl.write.Label(5, a, all.get(a).getWord()));
		}
		
		writeWorkBook.write();
		writeWorkBook.close();

	}
}