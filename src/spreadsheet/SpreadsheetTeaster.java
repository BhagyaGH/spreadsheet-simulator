package spreadsheet;
import java.util.*;

//Cell class definition
class Cell extends Observable implements Observer{
    private int rowNo;
    private int col;
    private String columnNo;
    private String cellAddress;
    private boolean isText;
    private String expression;
    private String cellText;
    private int cellValue;
    private Spreadsheet spreadsheet = new Spreadsheet();
    private NumberFunction num_Fun = new NumberFunction();
    private StringFunction string_Fun = new StringFunction();

    //Constructors for the Cell
    public Cell(String cellAddress, Spreadsheet spreadsheet){
        this.cellAddress = cellAddress;
        columnNo = cellAddress.substring(0, 0);
        rowNo = Integer.parseInt(cellAddress.substring(1));
        this.spreadsheet = spreadsheet;
    }
    public Cell(int column, int row){
        this.col = column;
        this.rowNo = row;
    }
    public Cell(int row, String column, Spreadsheet spreadsheet){
        rowNo = row;
        columnNo = column;
        cellAddress = columnNo + Integer.toString(rowNo);
        this.spreadsheet = spreadsheet;
    }
    //Set the cell details
    public void setCell(int row, String column, String value, boolean isText, Spreadsheet spreadsheet){
        rowNo = row;
        columnNo = column;
        cellAddress = columnNo + Integer.toString(rowNo);
        this.isText = isText;
        expression = value;
        this.spreadsheet = spreadsheet;
    }
    //Set the cell value
    public void setCellValue() {
        if(isText==true) {
            setText();
        }
        else {
            setValue();
        }
    }
    public String getCellAddress(){return cellAddress;}
    public int getRowNo(){return rowNo;}
    public String getColumnNo(){return columnNo;}
    public String getExpression(){return expression;}
    public boolean getIsText() {return isText;}
    //Set the numeric value of the cell
    public void setValue(){
        //If it's not an expression then store the number
        if(expression.charAt(0) != '='){
            cellValue = Integer.parseInt(expression);
        }
        //Evaluate the expression
        else{
            StringTokenizer input = new StringTokenizer(expression, "=(:)+-*/", true);
            int x = 0;
            int val_Int[] = new int[10];
            String operator[] = new String[5];
            String val[] = new String[input.countTokens()];
            Cell observableCell [] = new Cell[10];

                while(input.hasMoreTokens()) {
                    operator[x] = input.nextToken();
                    if(operator[x].equals(")")){
                        break;
                    }
                    val[x] = input.nextToken();
                    x++;
                }
                if(val[0].equals("SUM")||val[0].equals("sum")||val[0].equals("Sum")){
                    int ans=0;
                    int col1 = Character.getNumericValue(val[1].charAt(0)) - 10;
                    int row1 = Integer.parseInt(val[1].substring(1));
                    int col2 = Character.getNumericValue(val[2].charAt(0)) - 10;
                    int row2 = Integer.parseInt(val[2].substring(1));
                    int m = 0;
                    for(int j=col1; j<=col2; j++){
                        for(int i=row1; i<=row2; i++){
                            observableCell[m] = spreadsheet.getCell(i, j);
                            observableCell[m].addObserver(this);
                            m++;
                            ans = num_Fun.add(ans, spreadsheet.getCell(i, j).getValue());
                            this.setValue(ans);
                        }
                    }
                }
                else if(val[0].equals("MAX")||val[0].equals("max")||val[0].equals("Max")){
                    int col1 = Character.getNumericValue(val[1].charAt(0)) - 10;
                    int row1 = Integer.parseInt(val[1].substring(1));
                    int col2 = Character.getNumericValue(val[2].charAt(0)) - 10;
                    int row2 = Integer.parseInt(val[2].substring(1));
                    int max = spreadsheet.getCell(row1, col1).getValue();
                    int m = 0;
                    for(int j=col1; j<=col2; j++){
                        for(int i=row1+1; i<=row2; i++, m++){
                            observableCell[m] = spreadsheet.getCell(i, j);
                            observableCell[m].addObserver(this);
                            if(max < observableCell[m].getValue()) {
                                max = observableCell[m].getValue();
                            }
                            else {
                                continue;
                            }
                        }
                    }
                    this.setValue(max);
                }
                else{
                for(int j=0; j<2; j++){
                     try{
                         val_Int[j] = Integer.parseInt(val[j]);
                     }
                     catch (NumberFormatException n){
                        int col = Character.getNumericValue(val[j].charAt(0));
                        int row = Integer.parseInt(val[j].substring(1));
                        int m = 0;
                        observableCell[m] = spreadsheet.getCell(row, col-10);
                        observableCell[m].addObserver(this);
                        val_Int[j] = observableCell[m].getValue();
                        m++;
                     }
                }
                int cellVal_1 = val_Int[0];
                int cellVal_2 = val_Int[1];

                if(operator[1] == null){this.setValue(cellVal_1);}
                char opr = operator[1].charAt(0);
                switch(opr){
                     case '+': this.setValue(num_Fun.add(cellVal_1, cellVal_2)); break;
                     case '-': this.setValue(num_Fun.subtract(cellVal_1, cellVal_2)); break;
                     case '*': this.setValue(num_Fun.multiply(cellVal_1, cellVal_2)); break;
                     case '/': this.setValue(num_Fun.divide(cellVal_1, cellVal_2)); break;
                     default: break;
                }
                }
        }
    }
    public void setText(){
        cellText = expression;
    }
    public void setValue(int value){
        cellValue = value;
    }
    public int getValue(){return cellValue;}
    public String getText(){return cellText;}
    public void clear(){
        expression = null;
        cellValue = 0;
        cellText = null;
        setChanged();
        notifyObservers();
    }
    //Changing the cell value
    public void changeCellValue(String value, boolean isText){
        expression = value;
        this.isText = isText;
        if(isText) {
            cellText = value;
        }
        else {
            this.setValue();
        }
        setChanged();
        notifyObservers();
    }
    @Override
    public void update(Observable cell, Object change){
        if(!isText) {
            this.setValue();
        }
    }
}


    //Column class definition
    class Column {
        String columnNo;
        int noOfRows;
        Cell [] cells;
        int i = 0, j = 0;
        Spreadsheet spreadsheet = new Spreadsheet();

        //Constructor for Column
        public Column(String column, int rows, Spreadsheet spreadsheet){
            columnNo = column;
            noOfRows = rows;
            cells = new Cell[noOfRows];
            this.spreadsheet = spreadsheet;
        }
        public void setColumnNo(String column){columnNo = column;}
        public String getColumnNo(){return columnNo;}
        //Create the cells of the columns
        public void cellInitillize(){
            for(j=0; j<noOfRows; j++) {
                cells[j] = new Cell(j, columnNo, spreadsheet);
            }
        }
        //Store the details of the cells
        public void storeCell(String cellAddress, String value, boolean isText){
            String column = cellAddress.substring(0, 0);
            int row = Integer.parseInt(cellAddress.substring(1));
            cells[i].setCell(row, column, value, isText, spreadsheet);
            cells[i].setCellValue();
            if(i<j) {
                i++;
            }
        }
        public void storeCell(int row, String column, String value, boolean isText){
            cells[i].setCell(row, column, value, isText, spreadsheet);
            cells[i].setCellValue();
            if(i<j) {
                i++;
            }
        }
        //Return the cell in the given cell address
        public Cell getCell(int row){return cells[row];}
    }

    //NumberFunction class definition
    class NumberFunction{
        int answer;

        //Methods for arithmetic operations
        public int add(int x, int y){answer = x + y;return answer;}
        public int subtract(int x, int y){answer = x - y;return answer;}
        public int multiply(int x, int y){answer = x * y;return answer;}
        public int divide(int x, int y){answer = x / y;return answer;}
    }

    class StringFunction{
        String answer;

        public String cocantenate(String x, String y){
            answer = x + y;
            return answer;
        }
    }

//Class for spreadsheet
class Spreadsheet{
    int noOfColumns;
    Column columns [];
    int rows;
    int j = 0;
    int colStart, rowStart, colEnd, rowEnd;
    Cell [][] cellClipBoard;       

    public void setColumns(int columnsNo, int rows){
        noOfColumns = columnsNo;
        columns = new Column[noOfColumns];
        this.rows = rows;
    }
    public void addColumn(String name,Spreadsheet spreadsheet){
        columns[j] = new Column(name, rows, spreadsheet);
        columns[j].cellInitillize();
        j++;
    }
    public void storeCell(int row, String column, String value, boolean isText) { 
        int k;
        for(k=0; k<j; k++){
            if(columns[k].getColumnNo().equals(column)) {
                break;
            }
        }
        columns[k].storeCell(row, column, value, isText);
    }
    //Return the requested cell
    public Cell getCell(int row, int col) {
        return columns[col].getCell(row);
    }
    //Selecting a region of cells
    public void selectCells(String Address1, String Address2) {
        this.colStart = Character.getNumericValue(Address1.charAt(0)) - 10;
        this.rowStart = Integer.parseInt(Address1.substring(1));
        this.colEnd = Character.getNumericValue(Address2.charAt(0)) - 10;
        this.rowEnd = Integer.parseInt(Address2.substring(1));
    }
    //Storing the selected region to a clipboard without changing the original region
    public void copy() { 
        cellClipBoard = new Cell[rowEnd-rowStart+2][colEnd-colStart+2];
        for(int j=colStart; j<=colEnd; j++) {
            for(int i=rowStart; i<=rowEnd; i++) {    
                cellClipBoard[i][j] = new Cell(j, i);
                cellClipBoard[i][j] = columns[j].getCell(i);
            }
        }
    }
    //Storing the selected region to a clipboard and deleting the original region
    public void cut() {     
        cellClipBoard = new Cell[rowEnd-rowStart+2][colEnd-colStart+2];
        for(int i=rowStart; i<=rowEnd; i++) {
            for(int j=colStart; j<=colEnd; j++) { 
                cellClipBoard[i][j] = new Cell(j, i);
                cellClipBoard[i][j] = columns[j].getCell(i);
                columns[j].getCell(i).clear();
            }
        }
    }
    //Pasting the required cells in the clipboard in the given region
    public void paste(String Address3, String Address4) {
        int colStart2 = Character.getNumericValue(Address3.charAt(0)) - 10;
        int rowStart2 = Integer.parseInt(Address3.substring(1));
        int colEnd2 = Character.getNumericValue(Address4.charAt(0)) - 10;
        int rowEnd2 = Integer.parseInt(Address4.substring(1));  
        int i2 = rowStart2, j2 = colStart2; 
        for(int j=colStart; j<=colEnd; j++) {
            i2 = rowStart2;
            for(int i=rowStart; i<=rowEnd; i++) {
                columns[j2].getCell(i2).changeCellValue(cellClipBoard[i][j].getExpression(), cellClipBoard[i][j].getIsText());                        
                i2++;
            }   
            j2++;
        }
    } 
    //Deleting the selected cell region
    public void delete() {
        for(int i=rowStart; i<=rowEnd; i++) {
            for(int j=colStart; j<=colEnd; j++) { 
                columns[j].getCell(i).clear();
            }
        }
    }
}

//Class for testing the spreadsheet
public class SpreadsheetTeaster {

    public static void main(String[] args) {
       //For simulation purposes I make only 4 colums of 10 rows each for my Spreadsheet
       Spreadsheet spreadsheet1 = new Spreadsheet();
       System.out.println("--------------Spreadsheet Simulation--------------");
       System.out.printf("\nSpreadsheet consists of 4 Columns & 10 Rows\n");

       spreadsheet1.setColumns(4, 10);
       spreadsheet1.addColumn("A", spreadsheet1);
       spreadsheet1.addColumn("B", spreadsheet1);
       spreadsheet1.addColumn("C", spreadsheet1);
       spreadsheet1.addColumn("D", spreadsheet1);

       //Storing values to the cells
       spreadsheet1.storeCell(0, "A", "5", false);
       spreadsheet1.storeCell(1, "A", "=A0+10", false);
       spreadsheet1.storeCell(0, "B", "7", false);
       spreadsheet1.storeCell(1, "B", "1", false);
       spreadsheet1.storeCell(0, "C", "=A0-B0", false);
       spreadsheet1.storeCell(1, "C", "=9*B1", false);
       spreadsheet1.storeCell(2, "C", "=SUM(A0:B1)", false);
       spreadsheet1.storeCell(0, "D", "=MAX(A0:C1)", false);

       System.out.println("\nExpressions & Values stored in the cells\n");

       System.out.printf("Expression of %s = %s\t\tValue of %s = %d\n", spreadsheet1.getCell(0, 0).getCellAddress(), spreadsheet1.getCell(0, 0).getExpression(), spreadsheet1.getCell(0, 0).getCellAddress(), spreadsheet1.getCell(0, 0).getValue());
       System.out.printf("Expression of %s = %s\tValue of %s = %d\n", spreadsheet1.getCell(1, 0).getCellAddress(), spreadsheet1.getCell(1, 0).getExpression(), spreadsheet1.getCell(1, 0).getCellAddress(), spreadsheet1.getCell(1, 0).getValue());
       System.out.printf("Expression of %s = %s\t\tValue of %s = %d\n", spreadsheet1.getCell(0, 1).getCellAddress(), spreadsheet1.getCell(0, 1).getExpression(), spreadsheet1.getCell(0, 1).getCellAddress(), spreadsheet1.getCell(0, 1).getValue());
       System.out.printf("Expression of %s = %s\t\tValue of %s = %d\n", spreadsheet1.getCell(1, 1).getCellAddress(), spreadsheet1.getCell(1, 1).getExpression(), spreadsheet1.getCell(1, 1).getCellAddress(), spreadsheet1.getCell(1, 1).getValue());
       System.out.printf("Expression of %s = %s\tValue of %s = %d\n", spreadsheet1.getCell(0, 2).getCellAddress(), spreadsheet1.getCell(0, 2).getExpression(), spreadsheet1.getCell(0, 2).getCellAddress(), spreadsheet1.getCell(0, 2).getValue());
       System.out.printf("Expression of %s = %s\tValue of %s = %d\n", spreadsheet1.getCell(1, 2).getCellAddress(), spreadsheet1.getCell(1, 2).getExpression(), spreadsheet1.getCell(1, 2).getCellAddress(), spreadsheet1.getCell(1, 2).getValue());
       System.out.printf("Expression of %s = %s\tValue of %s = %d\n", spreadsheet1.getCell(2, 2).getCellAddress(), spreadsheet1.getCell(2, 2).getExpression(), spreadsheet1.getCell(2, 2).getCellAddress(), spreadsheet1.getCell(2, 2).getValue());
       System.out.printf("Expression of %s = %s\tValue of %s = %d\n", spreadsheet1.getCell(0, 3).getCellAddress(), spreadsheet1.getCell(0, 3).getExpression(), spreadsheet1.getCell(0, 3).getCellAddress(), spreadsheet1.getCell(0, 3).getValue());

       spreadsheet1.getCell(0, 0).changeCellValue("25", false);
       System.out.println("\nNow change the value of cell A0 to 25\nNow the values are changed\n");
       System.out.printf("Expression of %s = %s\t\tValue of %s = %d\n", spreadsheet1.getCell(0, 0).getCellAddress(), spreadsheet1.getCell(0, 0).getExpression(), spreadsheet1.getCell(0, 0).getCellAddress(), spreadsheet1.getCell(0, 0).getValue());
       System.out.printf("Expression of %s = %s\tValue of %s = %d\n", spreadsheet1.getCell(1, 0).getCellAddress(), spreadsheet1.getCell(1, 0).getExpression(), spreadsheet1.getCell(1, 0).getCellAddress(), spreadsheet1.getCell(1, 0).getValue());

       System.out.println("Selecting & copying the cells through A0:B1");
       spreadsheet1.selectCells("A0", "B1");
       spreadsheet1.copy();
       spreadsheet1.paste("C0", "D1");
      
       System.out.println("Pasting the cells through A0:B1 into C0:D1\nNew values stored are:\n");

       System.out.printf("Expression of %s = %s\t\tValue of %s = %d\n", spreadsheet1.getCell(0, 2).getCellAddress(), spreadsheet1.getCell(0, 2).getExpression(), spreadsheet1.getCell(0, 2).getCellAddress(), spreadsheet1.getCell(0, 2).getValue());
       System.out.printf("Expression of %s = %s\tValue of %s = %d\n", spreadsheet1.getCell(1, 2).getCellAddress(), spreadsheet1.getCell(1, 2).getExpression(), spreadsheet1.getCell(1, 2).getCellAddress(), spreadsheet1.getCell(1, 2).getValue());
       System.out.printf("Expression of %s = %s\t\tValue of %s = %d\n", spreadsheet1.getCell(0, 3).getCellAddress(), spreadsheet1.getCell(0, 3).getExpression(), spreadsheet1.getCell(0, 3).getCellAddress(), spreadsheet1.getCell(0, 3).getValue());
       System.out.printf("Expression of %s = %s\t\tValue of %s = %d\n", spreadsheet1.getCell(1, 3).getCellAddress(), spreadsheet1.getCell(1, 3).getExpression(), spreadsheet1.getCell(1, 3).getCellAddress(), spreadsheet1.getCell(1, 3).getValue());
       
       System.out.println("\nSelecting & deleting the cell D1");
       spreadsheet1.selectCells("D1", "D1");
       spreadsheet1.delete();
       System.out.println("Expression & value of cell D1 after deletion:\n");
       System.out.printf("Expression of %s = %s\t\tValue of %s = %d\n\n", spreadsheet1.getCell(1, 3).getCellAddress(), spreadsheet1.getCell(1, 3).getExpression(), spreadsheet1.getCell(1, 3).getCellAddress(), spreadsheet1.getCell(1, 3).getValue());
       
    }
}

