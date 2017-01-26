
public class ExcelFileInfo {
	
	public String fileName;
	public String location;
	public String sheet;
	public String cell;
	public TreeNode<Dependencies> dependentCell;
			
	public ExcelFileInfo() {
		fileName = null;
		location = null;
		sheet = null;
		cell = null;
		
		Dependencies root = new Dependencies(null, null);
		dependentCell = new TreeNode<Dependencies>(root);
	}
	
	public ExcelFileInfo(String filename, String location, String sheet, String cell) {
		this.fileName = filename;
		this.location = location;
		this.sheet = sheet;
		this.cell = cell;
		
		Dependencies root = new Dependencies(sheet, cell);
		dependentCell = new TreeNode<Dependencies>(root);
		
	}

}
