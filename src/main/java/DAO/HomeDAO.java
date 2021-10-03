package DAO;

import java.io.*;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.SQLException;

import javax.servlet.RequestDispatcher;
import javax.servlet.ServletException;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.catalina.tribes.ChannelSender;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import BEAN.Users;
import DB.DBConnection;


import BEAN.Users;

public class HomeDAO 
{
			public static void ImportExcel(HttpServletRequest request,HttpServletResponse response, Connection conn) throws ServletException, IOException
			{
				/*FileInputStream inp;
				try 
				{
					inp = new FileInputStream("F://test.xls");
					HSSFWorkbook wb = new HSSFWorkbook(new POIFSFileSystem(inp));
					
					Sheet sheet = wb.getSheetAt(0);
					
					
					
					for (int i=1; i<=sheet.getLastRowNum();i++)
					{
						Row row = sheet.getRow(i);
						
						int stt = (int) row.getCell(0).getNumericCellValue();
						
						String name = row.getCell(1).getStringCellValue();
						int pass = (int) row.getCell(2).getNumericCellValue();
						
						
						Users users = new Users();
						users.setStt(stt);
						users.setName(name);
						users.setPass(pass);
						
						HomeDAO.InsertData(request, users, conn);
						
					}
					
					wb.close();
					
				} 
				catch (FileNotFoundException e) 
				{
					request.setAttribute("message",e.getMessage());
					
				}
				catch (IOException e) 
				{
					request.setAttribute("message",e.getMessage());
					
				}
				
				RequestDispatcher rd = request.getRequestDispatcher("View/Result.jsp");
				rd.forward(request,response); "D://Book1.xls*/
				FileInputStream inp;
				try 
				{
					inp = new FileInputStream("D://Book1.xls");
					HSSFWorkbook wb = new HSSFWorkbook(new POIFSFileSystem(inp));
					
					Sheet sheet = wb.getSheetAt(0);
					
					int stt = 0;
					String name = " ";
					int pass = 0 ;
					for (int i=1; i<=sheet.getLastRowNum();i++)
					{
						Row row = sheet.getRow(i);
						if(row.getCell(0)==null) {
							stt = 0;
						}
						else if(row.getCell(1)==null) {
						name = " ";
						}
						else if(row.getCell(2)==null) {
							pass = 0;
						}
						else {
							 stt = (int) row.getCell(0).getNumericCellValue();
								
							 name = row.getCell(1).getStringCellValue();
							 pass = (int) row.getCell(2).getNumericCellValue();
							
							
						}
					
						
						
						Users users = new Users();
						users.setStt(stt);
						users.setName(name);
						users.setPass(pass);
						
						HomeDAO.InsertData(request, users, conn);
						
					}
					
					wb.close();
					
				} 
				catch (FileNotFoundException e) 
				{
					request.setAttribute("message",e.getMessage());
					
				}
				catch (IOException e) 
				{
					request.setAttribute("message",e.getMessage());
					
				}
				
				RequestDispatcher rd = request.getRequestDispatcher("/WEB-INF/View/Result.jsp");
				rd.forward(request,response);
			}
			
			public static void InsertData(HttpServletRequest request,Users users, Connection conn)
			{
				String sql = "insert into account(stt,name,pass) values (?,?,?)";
				try 
				{
					PreparedStatement ptmt = conn.prepareStatement(sql);
					
					
					ptmt.setInt(1,users.getStt());
					ptmt.setString(2,users.getName());
					ptmt.setInt(3,users.getPass());
					
					int kt = ptmt.executeUpdate();
					
					if (kt!=0)
					{
						request.setAttribute("message","Insert data from excel to mysql  success");
					}
					else 
					{
						request.setAttribute("message","Insert data from excel to mysql  failed");
					}
					
					ptmt.close();
					
				} 
				catch (SQLException e) 
				{				
					request.setAttribute("message",e.getMessage());
				}
			}
		    // Get cell value
		    private static Object getCellValue(Cell cell) {
		        CellType cellType = cell.getCellTypeEnum();
		        Object cellValue = null;
		        switch (cellType) {
		        case BOOLEAN:
		            cellValue = cell.getBooleanCellValue();
		            break;
		        case FORMULA:
		            Workbook workbook = cell.getSheet().getWorkbook();
		            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
		            cellValue = evaluator.evaluate(cell).getNumberValue();
		            break;
		        case NUMERIC:
		            cellValue = cell.getNumericCellValue();
		            break;
		        case STRING:
		            cellValue = cell.getStringCellValue();
		            break;
		        case _NONE:
		        case BLANK:
		        case ERROR:
		            break;
		        default:
		            break;
		        }
		 
		        return cellValue;
		    }

}
