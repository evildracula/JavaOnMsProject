package ext.project.ms.msproject;

import java.io.Serializable;

public interface IProjectData extends Serializable {

	public void update(String colName, Object colVal);

	public void delete();

}