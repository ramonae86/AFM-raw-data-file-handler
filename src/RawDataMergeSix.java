import java.io.*;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;


import javax.swing.JDialog;
import javax.swing.JOptionPane;

/*********************************************************************************************************
 * This program is dedicated to my dear, to help her get rid of the repetitive work on data processing.  *
 *                                                                                                       *
 * @author Ramon                                                                                         *
 *                                          ******   ******                                              *                   
 *                                       ********** **********                                           *
 *                                    ***************************                                        *
 *                                   *****************************                                       *
 *                                   *****************************                                       *
 *                                   *****************************                                       *
 *                                    ***************************                                        *
 *                                      ***********************                                          *
 *                                        *******************                                            *
 *                                          ***************                                              *
 *                                            ***********                                                *
 *                                              *******                                                  *
 *                                                ***                                                    *
 *                                                 *                                                     *
 *                                                                                                       *
 ********************************************************************************************************/

public class RawDataMergeSix
{
	/**
	 *@prefix prefix of filename
	 *@groupNum index of which group to process
	 *
	 *@targetBook WorkBook object to be written
	 *@fileOut excel file to be set up
	 *@sheetWrite set up a new sheet named Sheet 1 in targetBook
	 *@rowWrite set up a new row in targetBook
	 *@cellWrite set up a new cell in targetBook
	 *
	 *@rawDataArray a 1024x1 array that stores the data extracted from one data file
	 *@dataArray a 512x20 matrix that stores all the data for all ten data files
	 *@diffDataArray a 512x20 matrix that stores the differentiation of dataArray
	 *@largestDiffPosition a 20x1 array that stores the position where the largest differentiation lies
	 */
	 
	private static String prefix;
	private static int groupNum;
	
	private static HSSFWorkbook targetBook = new HSSFWorkbook();
	private static FileOutputStream fileOut;
	private static HSSFSheet sheetWrite = targetBook.createSheet("Sheet 1");
	private static HSSFRow rowWrite;
	private static HSSFCell cellWrite;
	
	private static int[] rawDataArray = new int[1024];
	private static int[][] dataArray = new int[512][10];
	private static int[][] diffDataArray = new int[512][10];
	private static int[] largestDiffPosition = new int[10];
	
	//two indexes
	private static int i = 1,j = 0;
	
	public static void main(String [] Args)
	{		
		try
		{
			//ask for file pre-fix
			prefix = JOptionPane.showInputDialog("What is the pre-fix of the raw data files?");
			//ask for the group number
			groupNum = Integer.parseInt(JOptionPane.showInputDialog("What is the group number?"));
			//set up a new file
			fileOut= new FileOutputStream(groupNum + "0-"+ groupNum +"4.xls");
			
			getData();
			
			//set font and style of the sheet
			HSSFFont font = targetBook.createFont();
			font.setFontName("����");
			font.setFontHeightInPoints((short)12);
			HSSFCellStyle style = targetBook.createCellStyle();
			style.setFont(font);
			//set the width of cell
			for(i = 0;i <= 30;i++)
				sheetWrite.setColumnWidth(i, 9 * 256);
		    
			//write dataArray values into Sheet 1
			for(i = 0;i <= 511;i++)
			{
				rowWrite = sheetWrite.createRow((short)(i));
				for(j = 0;j <= 9;j++)
				{
					cellWrite = rowWrite.createCell(j);
					cellWrite.setCellValue(dataArray[i][j]);
					cellWrite.setCellStyle(style);
				}
			}
			
/*			//generate the differentiation of dataArray
			for(j = 0;j <= 9;j++)
			{
				diffDataArray[0][j] = 0;
				for(i = 1;i <= 511;i++)
				{
					diffDataArray[i][j] = dataArray[i][j]-dataArray[i - 1][j];
				}
			}
			
			//find positions and store the row index numbers in an array
			int[] position = new int[10];
			for(j = 0;j <= 9;j++)
			{
				int largestDiff = 0;
				
				int extremeValuePosition = 0;
				for(i = 0;i <= 511;i++)
				{
					if(diffDataArray[i][j] >= largestDiff)
					{
						largestDiff = diffDataArray[i][j];
						largestDiffPosition[j] = i;
					}
				}
				
				//display the positions with largest differentiation
				System.out.println(largestDiffPosition[j]);
				
				//find local extreme value around the largest differentiation position
				int k = i;
				if(j%2 == 0)
				{
					for(k = 0;k < largestDiffPosition[j] + 10;k++)
					{
						if(diffDataArray[k][j] > 0)
						{
							extremeValuePosition = k;
							System.out.println(k + " " + j);
							break;
						}
					}
					position[j] = extremeValuePosition;
				}
				else
				{
					for(k = largestDiffPosition[j] - 5;k < 511;k++)
					{
						if(diffDataArray[k][j] < 0)
						{
							extremeValuePosition = k;
							System.out.println(k + " " + j);
							break;
						}
					}
					position[j] = extremeValuePosition;
				}
				
				//set the font for detected cells
				HSSFFont fontPosition = targetBook.createFont();
				fontPosition.setFontName("����");
				fontPosition.setFontHeightInPoints((short)12);
				HSSFCellStyle stylePosition = targetBook.createCellStyle();
				stylePosition.setFont(fontPosition);
				stylePosition.setFillForegroundColor(HSSFColor.YELLOW.index);
				stylePosition.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
				
				sheetWrite.getRow(position[j] - 1).getCell(j).setCellStyle(stylePosition);
			}*/
			
			targetBook.write(fileOut);
			fileOut.close();
			
			//generate txt file that records the positions
/*			PrintWriter resultTXT = new PrintWriter(groupNum + "0-"+ groupNum +"4.txt");
			for(j = 0;j <= 4;j++)
				resultTXT.println(position[2*j]);
			for(j = 0;j <= 4;j++)
				resultTXT.println(position[2*j + 1]);
			resultTXT.close();
*/
			
/*			
			//assign result string to positionResult			
			String positionResult = "Successfully done. Positions detected are:\n" + 
					position[0] + "                " + position[1] + "\n" + position[2] + "                " + position[3] + "\n" + 
					position[4] + "                " + position[5] + "\n" + position[6] + "                " + position[7] + "\n" + 
					position[8] + "                " + position[9];
			
			//display message dialog
			JOptionPane op = new JOptionPane(positionResult,JOptionPane.INFORMATION_MESSAGE);
			JDialog dialog = op.createDialog("Detected Positions");
			dialog.setAlwaysOnTop(true); //<-- this line
			dialog.setModal(true);
			dialog.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
			dialog.setVisible(true);
*/
		}
		catch(IOException e)
		{
			JOptionPane.showMessageDialog(null,"File doesn't exist.Program terminated." + e.getMessage());
		}
		catch(Exception e)
		{
			JOptionPane.showMessageDialog(null,"Error code:" + e.getMessage() + "\n" + e.toString() + "\n" + e.getStackTrace());
		}
	}
	//method that takes in an AFM file and generates a one dimension array with 1024 elements
	public int[] getRawData(String fileName)
	{
		try
		{
			DataInputStream inStream = new DataInputStream(new FileInputStream(fileName));
			/**
			 * @l the least significant byte
			 * @h the most significant byte
			 * @n order of the word read in
			 */
			int l;
			int dataStart = 40960;  //data starts from the 40960th word
			int n = 0;
			while((l = inStream.read()) != -1)
			{
				if(n < dataStart)
				{
					n += 1;
				}
				else //get data, store in rawDataArray
				{
					int h = inStream.read();
					System.out.println(l);
					System.out.println(h);
					//transfer UInt-8 Little Endian to decimal values
					if(256*h + l <= 32768)
						rawDataArray[(n-dataStart)/2] = 256*h + l;
					else
						rawDataArray[(n-dataStart)/2] = (256*h + l) - 65536;
					System.out.println(rawDataArray[(n-dataStart)/2]);
					n += 2;
				}
			}
			
			inStream.close();
			
			return rawDataArray;
		}
		catch(Exception e)
		{
			int[] array = new int[1];
			array[0] = -1;
			
			JOptionPane.showMessageDialog(null,"Error code:" + e.getMessage() + "\n" + e.toString() + "\n" + e.getStackTrace());
			
			return array;
		}
	}
	
	//get all the data from 6 raw data files and store the data in the 2-D array
	public static void getData()
	{
		//open files and move data from rawDataArray to dataArray
		for(int attempt = 0;attempt <= 99;attempt++)
		{
			//find all the files start with specified prefix
			String suffix;
			if(attempt == 0)
				suffix = ".000";
			else if(attempt <= 9)
				suffix = ".00" + attempt;
			else
				suffix = ".0" + attempt;
				
			File sourceBook = new File(prefix + "-" + groupNum + suffix);
			if(sourceBook.exists())
			{
				RawDataMergeSix object = new RawDataMergeSix();
				//store raw data in an array
				rawDataArray = object.getRawData(prefix + "-" + groupNum + suffix);
				//resize the ten 1024x1 arrays to twenty 512x1 arrays
				for(int k = 0;k <= 1023;k++)
				{
					if(k <= 511)
					{
						dataArray[k][2*j + 1] = rawDataArray[k];
//						System.out.println(dataArray[k][2*j + 1]);
//						System.out.println(dataArray[k][2*j + 3]);
					}
					else
					{
						dataArray[k - 512][2*j] = rawDataArray[k];
//						System.out.println(dataArray[k - 512][2*j]);
					}
				}
				j++;
			}
		}
	}
}
