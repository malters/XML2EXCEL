import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;

import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;


public class ReadXmlToExcelFile extends JFrame {
	public static int numNodo=0;
	public static ArrayList<String> estructuraXML = new ArrayList<String>();
	
	public static int FILAS = 0;
	public static int COLUMNAS = 0;
	public static Document doc;
	
	public ReadXmlToExcelFile(final NodeList listDeclaration) {
		JFrame fr=new JFrame();
		JPanel p=new JPanel();  
        final JCheckBox cb[]=new JCheckBox[COLUMNAS]; 
        JButton generar= new JButton("Generar EXCEL COLUMNAS SELECCIONADAS");
        JButton generarTODO= new JButton("Generar EXCEL TODAS COLUMNAS");
        
    	getNameColumnsXML(listDeclaration);
    	
        generar.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent ae) {
            	COLUMNAS=0;
            	estructuraXML = new ArrayList<String>();
            	for (int i = 0; i <cb.length; i++){
            	    if (cb[i].isSelected()){
            	    	COLUMNAS++;
            	    	estructuraXML.add(cb[i].getText());
            	    }
            	}
            	String[][] datosExport = new String[FILAS][COLUMNAS];	
            	addDataColumns(datosExport);
            	getDataColumnsXML(listDeclaration,datosExport,cb);
            	generateXLS(datosExport);
             }
        });
        
        generarTODO.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent ae) {
            	estructuraXML.clear();
            	numNodo=0;
            	getNameColumnsXML(listDeclaration);
                COLUMNAS=estructuraXML.size();
                
    			String[][] datosExport = new String[FILAS][COLUMNAS];	
    			
                addDataColumns(datosExport);        
                getDataColumnsXML(listDeclaration,datosExport);
                generateXLS(datosExport);
             }
        });
        
        for (int s = 0; s < estructuraXML.size(); s++) {
        	cb[s]=new JCheckBox(estructuraXML.get(s));
        	add(cb[s]);
        }
        
        for(int i=0;i<COLUMNAS;i++) {
            p.add(cb[i]);
          }
        p.add(generar);
        p.add(generarTODO);
        
        fr.setVisible(true);
        fr.setSize(400,400);
          
	    fr.add(p);
	 }
	 
	public static void main(String argv[]) {
		try{
	       
			DocumentBuilderFactory docBuilderFactory = DocumentBuilderFactory.newInstance();
	        DocumentBuilder docBuilder = docBuilderFactory.newDocumentBuilder();
	        doc = docBuilder.parse(new File("c:/TabletPcGps.wkt.xml"));
	        // normalize text representation
            doc.getDocumentElement().normalize();
            
            NodeList listDeclaration= doc.getElementsByTagName("LINEA_DECLARACION"); 
            int totalDeclaracion = listDeclaration.getLength();
            
            if (doc.hasChildNodes()) {

            	getNameColumnsXML(listDeclaration);

        	}
          
            FILAS=doc.getChildNodes().item(0).getChildNodes().getLength();
            COLUMNAS=estructuraXML.size();
            
            System.out.println("Columnas: " +COLUMNAS);
            System.out.println("Filas: " + FILAS);
  
			ReadXmlToExcelFile formulario1=new ReadXmlToExcelFile(listDeclaration);  
            
		}catch (Exception e) 
        {
            e.printStackTrace();
        }

        
	}
	
	private static void getNameColumnsXML(NodeList node){
        for (int s = 0; s < node.getLength(); s++) {
            Node nodo = node.item(s);  
            if (nodo.getNodeType() == Node.ELEMENT_NODE) {
            	Element eElementREGISTRO = (Element) nodo;
            	if(eElementREGISTRO.getNodeName().equalsIgnoreCase("LINEA_DECLARACION")){
            		numNodo++;
            	}
            	if(!eElementREGISTRO.getNodeName().equalsIgnoreCase("LINEA_DECLARACION") && numNodo<2){
            		estructuraXML.add(eElementREGISTRO.getNodeName());
            	}
            } 
           if (nodo.hasChildNodes() && numNodo<2) {
        	   getNameColumnsXML(nodo.getChildNodes());
           }
        }
	}
	
	private static String[][] addDataColumns(String[][]  datosExport){
        for (int s = 0; s < estructuraXML.size(); s++) {
            String cabecera = estructuraXML.get(s);
              datosExport[0][s]= cabecera;   	
        }
	  return datosExport;
	}
	
	private static String[][] getDataColumnsXML(NodeList node, String[][]  datosExport){
        for (int s = 0; s < node.getLength(); s++) {
            Node nodo = node.item(s);  
            if (nodo.getNodeType() == Node.ELEMENT_NODE) {
            	Element eElementREGISTRO = (Element) nodo;
                for (int a = 0; a < eElementREGISTRO.getChildNodes().getLength(); a++) {
                	Node nodoN = eElementREGISTRO.getChildNodes().item(a); 
                	datosExport[s+1][a]= nodoN.getTextContent();     	
                }
            }
        }
        
        return datosExport;
	}
	
	private static String[][] getDataColumnsXML(NodeList node, String[][]  datosExport,JCheckBox cb[]){
		int column=0;
        for (int s = 0; s < node.getLength(); s++) {
            Node nodo = node.item(s);  
            column=0;
            if (nodo.getNodeType() == Node.ELEMENT_NODE) {
            	Element eElementREGISTRO = (Element) nodo;
                for (int a = 0; a < eElementREGISTRO.getChildNodes().getLength(); a++) {
                	Node nodoN = eElementREGISTRO.getChildNodes().item(a); 
                	for (int i = 0; i <cb.length; i++){
                	    if (cb[i].isSelected() && nodoN.getNodeName().equalsIgnoreCase(cb[i].getText())){
                	    	datosExport[s+1][column]= nodoN.getTextContent();  
                	    	column++;
                	    }
                	    
                	}
                }
            }
        }
        
        return datosExport;
	}
	
	private static void generateXLS(String[][] datosExport){
		
		 XSSFWorkbook workbook = new XSSFWorkbook();
         XSSFSheet sheet = workbook.createSheet("Sample sheet");

         Map<String, Object[]> data = new HashMap<String, Object[]>();
         
         for(int i=0;i<1;i++)
         {
        	 Object[] columns = new Object[COLUMNAS];
        	 
             for(int a=0;a<COLUMNAS;a++)
             {
            	 columns[a]=datosExport[i][a];
             }
             data.put(i+"",columns);
         }
         
         Set<String> keyset = data.keySet();
         int rownum = 0;
         for (String key : keyset) {
             Row row = sheet.createRow(rownum++);
             Object[] objArr = data.get(key);
             int cellnum = 0;
             for (Object obj : objArr) {
                 Cell cell = row.createCell(cellnum++);
                 if (obj instanceof Date)
                     cell.setCellValue((Date) obj);
                 else if (obj instanceof Boolean)
                     cell.setCellValue((Boolean) obj);
                 else if (obj instanceof String)
                     cell.setCellValue((String) obj);
                 else if (obj instanceof Double)
                     cell.setCellValue((Double) obj);
             }
         }
         data = new HashMap<String, Object[]>();
         for(int i=1;i<datosExport.length;i++)
         {
        	 Object[] columns = new Object[COLUMNAS];
        	 
             for(int a=0;a<COLUMNAS;a++)
             {
            	 columns[a]=datosExport[i][a];
             }
             data.put(i+"",columns);
         }
         
          keyset = data.keySet();
         for (String key : keyset) {
             Row row = sheet.createRow(rownum++);
             Object[] objArr = data.get(key);
             int cellnum = 0;
             for (Object obj : objArr) {
                 Cell cell = row.createCell(cellnum++);
                 if (obj instanceof Date)
                     cell.setCellValue((Date) obj);
                 else if (obj instanceof Boolean)
                     cell.setCellValue((Boolean) obj);
                 else if (obj instanceof String)
                     cell.setCellValue((String) obj);
                 else if (obj instanceof Double)
                     cell.setCellValue((Double) obj);
             }
         }
         
         try {
             FileOutputStream out = new FileOutputStream(new File("c:/book.xlsx"));
             workbook.write(out);
             out.close();
             
             JOptionPane.showMessageDialog(new JFrame(),"Excel creada con éxito","Dialog",JOptionPane.INFORMATION_MESSAGE);

         } catch (FileNotFoundException e) {
             e.printStackTrace();
         } catch (IOException e) {
             e.printStackTrace();
         }
		
	}
}
