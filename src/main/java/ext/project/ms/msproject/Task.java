package ext.project.ms.msproject;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class Task implements IProjectData {

	private static final long serialVersionUID = -5813764880353780153L;
	private Dispatch task;
	private ActiveXComponent msProjApp;
	private int taskID;

	public int getTaskID() {
		return taskID;
	}

	public void setTaskID(int taskID) {
		this.taskID = taskID;
	}

	public Dispatch getTask() {
		return task;
	}

	public void setTask(Dispatch task) {
		this.task = task;
	}

	public ActiveXComponent getMsProjApp() {
		return msProjApp;
	}

	public void setMsProjApp(ActiveXComponent msProjApp) {
		this.msProjApp = msProjApp;
	}

	public Task(Dispatch task, ActiveXComponent msProjApp) {
		this.task = task;
		this.msProjApp = msProjApp;
		this.taskID = Dispatch.get(task, "ID").getInt();
	}

	public void update(String colName, Object colVal) {
		int fieldID = Dispatch.invoke(msProjApp, "FieldNameToFieldConstant", Dispatch.Method,
				new Object[] { colName, COMObjectTypeConstant.pjTask }, new int[] { 1 }).getInt();
		Dispatch.invoke(task, "SetField", Dispatch.Method, new Object[] { fieldID, colVal }, new int[] { 1 });
	}

	public void delete() {
		Dispatch.invoke(task, "Delete", Dispatch.Method, new Object[] {}, new int[] { 1 });
	}

	public void markAsMileStone() {
		Dispatch.put(task, "MileStone", new Variant(true));
	}

	public void configPredecessors(String predecessors) {
		Dispatch.put(task, "Predecessors", predecessors);
	}

	public Assignment addAssignment(Resource resc) throws Exception {
		int rescId = Dispatch.get(resc.getResource(), "ID").getInt();
		Assignment assignment = null;
		Dispatch assignments = Dispatch.get(task, "Assignments").toDispatch();
		// do validation first
		int assCount = Dispatch.get(assignments, "Count").getInt();
		for (int i = 0; i < assCount; i++) {
			Dispatch assItDispatch = Dispatch
					.invoke(assignments, "Item", Dispatch.Get, new Object[] { i + 1 }, new int[] { 1 }).getDispatch();
			if (assItDispatch != null) {
				int assItRescID = Dispatch.get(assItDispatch, "ResourceID").getInt();
				if (assItRescID == rescId) {
					throw new Exception("Adding an existed resource");
				}
			}
		}
		Dispatch assignmentPatch = Dispatch
				.invoke(assignments, "Add", Dispatch.Method, new Object[] { taskID, rescId }, new int[] { 1 })
				.getDispatch();
		assignment = new Assignment(this, msProjApp, assignmentPatch);
		return assignment;
	}

}