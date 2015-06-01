package nvg.mm.td;

public class ModelNameRow {
	private String modelName;
	private int row;
	
	public ModelNameRow () {
		modelName = null;
		row = 0;
	}
	public void setModelName(String anyModel){
		modelName = anyModel;
	}
	public void setModelRow(int anyRow){
		row = anyRow;
	}
	public String getModelName(){
		return modelName;
	}
	public int getModelRow(){
		return row;
	}
		
}
