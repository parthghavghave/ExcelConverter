package main.resource;

import javax.swing.filechooser.FileSystemView;

public class constants {

	static FileSystemView fileSystemView = FileSystemView.getFileSystemView();
	public static String desktopPath = fileSystemView.getHomeDirectory().getAbsolutePath();
	public static String OUTPPUT_FILE_LOCATION = desktopPath+"\\output.xls";
	
	public static int SN_COLUMN_INDEX = 0;
	public static int FUP_COLUMN_INDEX = 1;
	public static int DOC_COLUMN_INDEX = 2;
	public static int MODE_COLUMN_INDEX = 3;
	public static int POLICYNO_COLUMN_INDEX = 4;
	public static int PREMIUM_COLUMN_INDEX = 5;
	public static int NAME_COLUMN_INDEX = 6;
	public static int NAMETWO_COLUMN_INDEX = 7;
	
	public static int SN_COLUMN_WIDTH = 1200;
	public static int FUP_COLUMN_WIDTH = 2000;
	public static int DOC_COLUMN_WIDTH = 3000;
	public static int MODE_COLUMN_WIDTH = 700;
	public static int POLICYNO_COLUMN_WIDTH = 3000;
	public static int PREMIUM_COLUMN_WIDTH = 3500;
	public static int NAME_COLUMN_WIDTH = 4500;
	public static int NAMETWO_COLUMN_WIDTH = 4500;
}
