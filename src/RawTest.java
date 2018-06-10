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

public class RawTest
{
	//prefix of filename
	private static String prefix;
	//index of which group to process
	private static int groupNum;
	
	//WorkBook object to be written
	private static HSSFWorkbook targetBook = new HSSFWorkbook();
	//excel file to be set up
	private static FileOutputStream fileOut;
	//set up a new sheet named Sheet 1
	private static HSSFSheet sheetWrite = targetBook.createSheet("Sheet 1");
	//set up a new row of the write file
	private static HSSFRow rowWrite;
	//set up a new cell of the write file
	private static HSSFCell cellWrite;
	
	//store raw data in an array
	private static int[] rawDataArray = new int[1024];
	//create an two-dimension 512 x 20 array to store data
	private static int[][] dataArray = new int[512][20];
	//the differentiation of dataArray
	private static int[][] diffDataArray = new int[512][20];;
	//array to store the position that has the largest differentiation
	private static int[] largestDiffPosition = new int[20];
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
			fileOut= new FileOutputStream(groupNum + "0-"+ groupNum +"9.xls");
			
			getData();
			
			//set font and style of the sheet
			HSSFFont font = targetBook.createFont();
			font.setFontName("ËÎÌו");
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
				for(j = 0;j <= 19;j++)
				{
					cellWrite = rowWrite.createCell(j);
					cellWrite.setCellValue(dataArray[i][j]);
					cellWrite.setCellStyle(style);
				}
			}
			
			//generate the differentiation of dataArray
			for(j = 0;j <= 19;j++)
			{
				diffDataArray[0][j] = 0;
				for(i = 1;i <= 511;i++)
				{
					diffDataArray[i][j] = dataArray[i][j]-dataArray[i - 1][j];
				}
			}
			
			//find positions and store the row index numbers in an array
			int[] position = new int[20];
			for(j = 0;j <= 19;j++)
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
				fontPosition.setFontName("ËÎÌו");
				fontPosition.setFontHeightInPoints((short)12);
				HSSFCellStyle stylePosition = targetBook.createCellStyle();
				stylePosition.setFont(fontPosition);
				stylePosition.setFillForegroundColor(HSSFColor.YELLOW.index);
				stylePosition.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
				
				sheetWrite.getRow(position[j] - 1).getCell(j).setCellStyle(stylePosition);
			}
			
			targetBook.write(fileOut);
			fileOut.close();
			
			//generate txt file that records the positions
			PrintWriter resultTXT = new PrintWriter(groupNum + "0-"+ groupNum +"9.txt");
			for(j = 0;j <= 9;j++)
				resultTXT.println(position[2*j]);
			for(j = 0;j <= 9;j++)
				resultTXT.println(position[2*j + 1]);
			resultTXT.close();
			
			//assign result string to positionResult			
			String positionResult = "Successfully done. Positions detected are:\n" + 
					position[0] + "                " + position[1] + "\n" + position[2] + "                " + position[3] + "\n" + 
					position[4] + "                " + position[5] + "\n" + position[6] + "                " + position[7] + "\n" + 
					position[8] + "                " + position[9] + "\n" + position[10] + "                " + position[11] + "\n" + 
					position[12] + "                " + position[13] + "\n" + position[14] + "                " + position[15] + "\n" + 
					position[16] + "                " + position[17] + "\n" + position[18] + "                " + position[19];
			
			//display message dialog
			JOptionPane op = new JOptionPane(positionResult,JOptionPane.INFORMATION_MESSAGE);
			JDialog dialog = op.createDialog("Detected Positions");
			dialog.setAlwaysOnTop(true); //<-- this line
			dialog.setModal(true);
			dialog.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
			dialog.setVisible(true);
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
			int n = 0;
			
			while((l = inStream.read()) != -1)
			{
				//data starts from the 8192th word
				if(n >= 8192)
				{
					int h = inStream.read();
					//transfer UInt-8 Little Endian to decimal values
					if(256*h + l <= 32768)
						rawDataArray[(n-8192)/2] = 256*h + l;
					else
						rawDataArray[(n-8192)/2] = (256*h + l) - 65536;
					System.out.println(rawDataArray[(n-8192)/2]);
					n += 2;
				}
				else
					n++;
			}
			
			System.out.println(l);
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
				RawDataMerge object = new RawDataMerge();
				//store raw data in an array
				rawDataArray = object.getRawData(prefix + "-" + groupNum + suffix);
				//resize the ten 1024x1 arrays to twenty 512x1 arrays
				for(int k = 0;k <= 1023;k++)
				{
					if(k <= 511)
					{
						dataArray[k][2*j + 1] = rawDataArray[k];
//						System.out.println(dataArray[k][2*j + 1]);
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
