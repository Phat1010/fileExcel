package DB;


import java.sql.*;


public class DBConnection {
	public static Connection CreateConnection(){
		Connection conn = null;
	//	conn = DriverManager.getConnection(url,username,password);
		String url = "jdbc:mysql://localhost:3306/users";
		String username ="root";
		String password = "phat";

		try {
			Class.forName("com.mysql.jdbc.Driver");
		
			//kich hoat drive
			//sau khi load xong
			
		} catch (ClassNotFoundException e) {
			//chua coppu file trong lib thi`
			e.printStackTrace();
			//xuat ra tb ko load dc drive
		}
		
		//create Connection
		try {
		
			conn = DriverManager.getConnection(url,username,password);
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return conn;	
		
	}
}
